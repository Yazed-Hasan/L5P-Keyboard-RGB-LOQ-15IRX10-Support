use std::{
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use rand::Rng;

use crate::{
    enums::{Effects, StarsPalette, StarsParams},
    manager::{effects::lamps, Inner},
};

use super::scene;

struct Shooting {
    x: f32,
    dir: f32,
    hue: f32,
}

pub fn play(manager: &mut Inner, start: StarsParams, rgb0: [u8; 12]) {
    let mut params = start.normalized();
    let mut rgb = rgb0;
    let mut phases = Vec::new();
    let mut hues = Vec::new();
    let mut active = Vec::new();
    let mut shooting: Option<Shooting> = None;
    let mut last = Instant::now();
    let mut last_density = params.density;
    let mut rng = rand::rng();

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::Stars { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        let n = manager.lamp_n();
        if phases.len() != n {
            phases = (0..n).map(|_| rng.random_range(0.0..std::f32::consts::TAU)).collect();
            hues = (0..n).map(|_| rng.random_range(0.0..360.0)).collect();
            active = vec![false; n];
            pick_stars(&mut active, params.density, &mut rng);
            last_density = params.density;
        } else if (params.density - last_density).abs() > 0.02 {
            pick_stars(&mut active, params.density, &mut rng);
            last_density = params.density;
        }

        let twinkle = params.twinkle;
        for phase in phases.iter_mut() {
            *phase += dt * (0.55 + twinkle * 1.8);
        }

        let drift = params.hue_drift * 28.0 * dt;
        for hue in hues.iter_mut() {
            *hue = (*hue + drift).rem_euclid(360.0);
        }

        if shooting.is_none() && params.shooting > 0.001 && rng.random::<f32>() < params.shooting * dt * 1.4 {
            let right = rng.random_bool(0.5);
            shooting = Some(Shooting {
                x: if right { -0.08 } else { 1.08 },
                dir: if right { 1.0 } else { -1.0 },
                hue: rng.random_range(0.0..360.0),
            });
        }

        let mut energy = vec![params.background; n];
        let width = (0.18 + params.size * 1.4).max(0.12);
        for i in 0..n {
            if !active[i] {
                continue;
            }
            let tw = ((phases[i].sin() * 0.5 + 0.5).powf(1.55) * 0.88 + 0.12).clamp(0.0, 1.0);
            splat(&mut energy, i, tw, width, n);
        }

        if let Some(shot) = shooting.as_mut() {
            shot.x += shot.dir * dt * (0.55 + params.twinkle * 0.9);
            let head = (shot.x * (n.saturating_sub(1) as f32)).clamp(-2.0, n as f32 + 2.0);
            for i in 0..n {
                let dist = (i as f32 - head).abs();
                let trail = if shot.dir > 0.0 { (head - i as f32).max(0.0) } else { (i as f32 - head).max(0.0) };
                let core = (-dist * dist / 0.55).exp();
                let tail = (-trail * 1.6).exp() * 0.55;
                energy[i] = (energy[i] + (core + tail) * 0.95).min(1.0);
            }
            if shot.x < -0.2 || shot.x > 1.2 {
                shooting = None;
            }
        }

        let mut lamps_now = vec![[0u8; 3]; n];
        for i in 0..n {
            let t = lamps::pos(i, n);
            let color = star_color(params.palette, rgb, t, hues[i], shooting.as_ref().map(|s| s.hue));
            lamps_now[i] = lamps::scale_rgb(color, energy[i]);
        }
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(33));
    }
}

fn pick_stars(active: &mut [bool], density: f32, rng: &mut impl Rng) {
    let n = active.len();
    let want = ((n as f32 * density).round() as usize).clamp(1, n);
    for slot in active.iter_mut() {
        *slot = false;
    }
    let mut left: Vec<usize> = (0..n).collect();
    for _ in 0..want {
        if left.is_empty() {
            break;
        }
        let idx = rng.random_range(0..left.len());
        active[left.swap_remove(idx)] = true;
    }
}

fn splat(energy: &mut [f32], center: usize, amt: f32, width: f32, n: usize) {
    let span = ((width * n as f32).ceil() as isize + 1).max(1);
    for d in -span..=span {
        let i = center as isize + d;
        if i < 0 || i >= n as isize {
            continue;
        }
        let dist = d as f32;
        let w = (-dist * dist / (2.0 * width * width * (n as f32 * 0.22).max(0.35))).exp();
        let i = i as usize;
        energy[i] = (energy[i] + amt * w).min(1.0);
    }
}

fn star_color(palette: StarsPalette, rgb: [u8; 12], t: f32, hue: f32, shoot_hue: Option<f32>) -> [u8; 3] {
    match palette {
        StarsPalette::Custom => scene::custom(&rgb, t),
        StarsPalette::White => [255, 244, 230],
        StarsPalette::Gold => [255, 196, 72],
        StarsPalette::Rainbow => scene::hsv(hue, 0.55, 1.0),
        StarsPalette::Random => scene::hsv(shoot_hue.unwrap_or(hue), 0.42, 1.0),
    }
}
