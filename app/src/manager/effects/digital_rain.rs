use std::{
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use rand::Rng;

use crate::{
    enums::{DigitalRainPalette, DigitalRainParams, Effects},
    manager::{effects::lamps, Inner},
};

use super::scene;

struct Drop {
    y: f32,
    speed: f32,
}

pub fn play(manager: &mut Inner, start: DigitalRainParams, rgb0: [u8; 12]) {
    legion_rgb_driver::debug_log("EFFECT: digital rain");
    let mut params = start.normalized();
    let mut rgb = rgb0;
    let mut last = Instant::now();
    let mut drops: Vec<Drop> = Vec::new();
    let mut rng = rand::rng();

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::DigitalRain { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        let n = manager.lamp_n().max(1);
        let want = (1.0 + params.density * 7.0).round() as usize;

        if drops.len() < want && rng.random::<f32>() < params.density * dt * 8.0 {
            drops.push(Drop {
                y: -0.2 - rng.random::<f32>() * 0.35,
                speed: 0.55 + params.speed * (0.7 + rng.random::<f32>() * 0.7),
            });
        }

        for drop in drops.iter_mut() {
            drop.y += drop.speed * dt;
        }
        drops.retain(|d| d.y < 1.35);

        let mut energy = vec![params.background; n];
        let trail = 0.12 + params.trail * 0.55;
        let head_w = if n > 4 { 0.06 } else { 0.18 };
        for drop in &drops {
            for i in 0..n {
                let t = lamps::pos(i, n);
                let dist = t - drop.y;
                let head = if dist.abs() < head_w {
                    1.0
                } else if dist < 0.0 && dist > -trail {
                    (1.0 + dist / trail).powf(1.35)
                } else {
                    0.0
                };
                energy[i] = energy[i].max(head);
            }
        }

        let lamps_now: Vec<[u8; 3]> = (0..n)
            .map(|i| {
                let t = lamps::pos(i, n);
                let e = energy[i].clamp(0.0, 1.0);
                lamps::scale_rgb(rain_color(params.palette, rgb, t, e), e)
            })
            .collect();
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(33));
    }
}

fn rain_color(palette: DigitalRainPalette, rgb: [u8; 12], t: f32, energy: f32) -> [u8; 3] {
    match palette {
        DigitalRainPalette::Matrix => {
            let h = 118.0 + energy * 18.0;
            let s = (0.95 - energy * 0.45).clamp(0.25, 1.0);
            scene::hsv(h, s, 0.2 + energy * 0.8)
        }
        DigitalRainPalette::Ice => scene::hsv(188.0, 0.55 - energy * 0.35, 0.25 + energy * 0.75),
        DigitalRainPalette::Custom => {
            let base = scene::custom(&rgb, t);
            lamps::mix_rgb(lamps::scale_rgb(base, 0.2), lamps::mix_rgb(base, [230, 255, 230], energy * 0.4), energy)
        }
    }
}
