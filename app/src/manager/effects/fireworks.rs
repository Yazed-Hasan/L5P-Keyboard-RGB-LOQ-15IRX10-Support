use std::{
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use rand::Rng;

use crate::{
    enums::{Effects, FireworksPalette, FireworksParams},
    manager::{effects::lamps, Inner},
};

use super::scene;

struct Burst {
    x: f32,
    age: f32,
    life: f32,
    hue: f32,
}

pub fn play(manager: &mut Inner, start: FireworksParams, rgb0: [u8; 12]) {
    legion_rgb_driver::debug_log("EFFECT: fireworks");
    let mut params = start.normalized();
    let mut rgb = rgb0;
    let mut last = Instant::now();
    let mut bursts: Vec<Burst> = Vec::new();
    let mut rng = rand::rng();

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::Fireworks { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        let n = manager.lamp_n();
        let spawn_p = params.rate * dt * 3.4;
        if bursts.len() < 10 && rng.random::<f32>() < spawn_p {
            bursts.push(Burst {
                x: rng.random::<f32>(),
                age: 0.0,
                life: 0.35 + params.trail * 0.85,
                hue: rng.random::<f32>() * 360.0,
            });
        }
        for burst in bursts.iter_mut() {
            burst.age += dt;
        }
        bursts.retain(|b| b.age < b.life);

        let mut energy = vec![params.background; n];
        let mut hues = vec![200.0f32; n];
        let radius = 0.08 + params.size * 0.42;
        for burst in &bursts {
            let fade = (1.0 - burst.age / burst.life).clamp(0.0, 1.0).powf(1.15);
            let pop = (burst.age / (burst.life * 0.22)).min(1.0);
            let r = radius * (0.35 + pop * 0.65);
            for i in 0..n {
                let t = lamps::pos(i, n);
                let dist = (t - burst.x).abs();
                let ring = (-((dist - r * 0.35) * (dist - r * 0.35)) / (2.0 * (r * 0.55 + 0.02).powi(2))).exp();
                let amt = (ring * fade).clamp(0.0, 1.0);
                if amt > energy[i] {
                    energy[i] = amt;
                    hues[i] = burst.hue;
                }
            }
        }

        let lamps_now: Vec<[u8; 3]> = (0..n)
            .map(|i| {
                let t = lamps::pos(i, n);
                let e = energy[i].clamp(0.0, 1.0);
                lamps::scale_rgb(burst_color(params.palette, rgb, t, hues[i], e), e)
            })
            .collect();
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(33));
    }
}

fn burst_color(palette: FireworksPalette, rgb: [u8; 12], t: f32, hue: f32, energy: f32) -> [u8; 3] {
    match palette {
        FireworksPalette::Festival => {
            let s = (0.9 - energy * 0.35).clamp(0.4, 1.0);
            scene::hsv(hue, s, 0.25 + energy * 0.75)
        }
        FireworksPalette::Ice => scene::hsv(190.0 + hue * 0.08, 0.4, 0.3 + energy * 0.7),
        FireworksPalette::Custom => {
            let base = scene::custom(&rgb, t);
            lamps::mix_rgb(base, [255, 240, 210], energy * 0.45)
        }
    }
}
