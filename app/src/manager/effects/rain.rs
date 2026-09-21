use std::{
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use rand::Rng;

use crate::{
    enums::{Direction, Effects, RainPalette, RainParams},
    manager::{effects::lamps, Inner},
};

use super::scene;

struct Drop {
    x: f32,
    v: f32,
    hue: f32,
}

pub fn play(manager: &mut Inner, start: RainParams, rgb0: [u8; 12]) {
    let mut params = start.normalized();
    let mut rgb = rgb0;
    let mut drops: Vec<Drop> = Vec::new();
    let mut trail = Vec::new();
    let mut splash = Vec::new();
    let mut last = Instant::now();
    let mut rng = rand::rng();

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::Rain { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        let n = manager.lamp_n();
        if trail.len() != n {
            trail = vec![0.0f32; n];
            splash = vec![0.0f32; n];
        }

        let right = matches!(params.direction, Direction::Right);
        let max_drops = ((2.0 + params.density * n as f32 * 1.4).round() as usize).clamp(2, n * 3);
        let spawn_rate = params.density * (2.2 + n as f32 * 0.08);
        if drops.len() < max_drops && rng.random::<f32>() < spawn_rate * dt {
            let jitter = (rng.random::<f32>() - 0.5) * params.wind * 0.25;
            drops.push(Drop {
                x: if right { -0.08 + jitter.abs() } else { 1.08 - jitter.abs() },
                v: (0.55 + rng.random::<f32>() * 0.7) * params.speed * (1.0 + params.wind * 0.4),
                hue: rng.random_range(0.0..360.0),
            });
        }

        let decay = (0.55 + (1.0 - params.wet) * 3.2 + (1.0 - params.trail) * 1.4) * dt;
        for slot in trail.iter_mut() {
            *slot = (*slot - decay).max(0.0);
        }
        for slot in splash.iter_mut() {
            *slot = (*slot - dt * 3.4).max(0.0);
        }

        let mut i = 0;
        while i < drops.len() {
            let dir = if right { 1.0 } else { -1.0 };
            drops[i].x += dir * drops[i].v * dt;
            let ended = if right { drops[i].x > 1.04 } else { drops[i].x < -0.04 };
            if ended {
                let edge = if right { n - 1 } else { 0 };
                splash[edge] = (splash[edge] + params.splash).min(1.4);
                drops.swap_remove(i);
                continue;
            }
            let head = drops[i].x.clamp(0.0, 1.0) * n.saturating_sub(1) as f32;
            let lamp = head.round() as usize;
            if lamp < n {
                trail[lamp] = trail[lamp].max(1.0);
                let behind = if right { -1 } else { 1 };
                let tlen = (1 + (params.trail * n as f32 * 0.45).round() as isize).max(1);
                for k in 1..=tlen {
                    let j = lamp as isize + behind * k;
                    if j < 0 || j >= n as isize {
                        break;
                    }
                    let fade = 1.0 - k as f32 / (tlen as f32 + 0.4);
                    let j = j as usize;
                    trail[j] = trail[j].max(fade * (0.35 + params.trail * 0.65));
                }
            }
            i += 1;
        }

        let mut lamps_now = vec![[0u8; 3]; n];
        for i in 0..n {
            let t = lamps::pos(i, n);
            let e = (trail[i] + splash[i] * 1.15).clamp(0.0, 1.0);
            if e <= 0.001 {
                continue;
            }
            let hue = drops
                .iter()
                .min_by(|a, b| {
                    let da = (a.x - t).abs();
                    let db = (b.x - t).abs();
                    da.partial_cmp(&db).unwrap_or(std::cmp::Ordering::Equal)
                })
                .map(|d| d.hue)
                .unwrap_or(200.0);
            lamps_now[i] = lamps::scale_rgb(rain_color(params.palette, rgb, t, hue), e);
        }
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(33));
    }
}

fn rain_color(palette: RainPalette, rgb: [u8; 12], t: f32, hue: f32) -> [u8; 3] {
    match palette {
        RainPalette::Custom => scene::custom(&rgb, t),
        RainPalette::Ice => scene::hsv(188.0 + t * 28.0, 0.55, 1.0),
        RainPalette::Neon => scene::hsv(132.0 + t * 40.0, 0.85, 1.0),
        RainPalette::Rainbow => scene::hsv(hue, 0.8, 1.0),
    }
}
