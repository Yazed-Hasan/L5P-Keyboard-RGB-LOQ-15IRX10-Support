use std::{
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use crate::{
    enums::{Direction, Effects, ScannerPalette, ScannerParams, ScannerPath},
    manager::{effects::lamps, Inner},
};

use super::scene;

pub fn play(manager: &mut Inner, start: ScannerParams, rgb0: [u8; 12]) {
    let mut params = start.normalized();
    let mut rgb = rgb0;
    let mut x = if matches!(start.direction, Direction::Left) { 1.0 } else { 0.0 };
    let mut dir = if matches!(start.direction, Direction::Left) { -1.0 } else { 1.0 };
    let mut trail = Vec::new();
    let mut last = Instant::now();
    let mut hue_phase = 0.0f32;

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::Scanner { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        hue_phase = (hue_phase + dt * 40.0).rem_euclid(360.0);
        let n = manager.lamp_n();
        if trail.len() != n {
            trail = vec![0.0f32; n];
        }

        let speed = 0.22 + params.speed * 0.55;
        x += dir * speed * dt;
        match params.path {
            ScannerPath::Bounce => {
                if x <= 0.0 {
                    x = 0.0;
                    dir = 1.0;
                } else if x >= 1.0 {
                    x = 1.0;
                    dir = -1.0;
                }
            }
            ScannerPath::Wrap => {
                if x < 0.0 {
                    x += 1.0;
                } else if x > 1.0 {
                    x -= 1.0;
                }
                if matches!(params.direction, Direction::Left) {
                    dir = -1.0;
                } else {
                    dir = 1.0;
                }
            }
        }

        let decay = (1.6 + (1.0 - params.trail) * 4.5) * dt;
        for slot in trail.iter_mut() {
            *slot = (*slot - decay).max(0.0);
        }

        let width = 0.035 + params.width * 0.16;
        let mut beam = vec![0.0f32; n];
        stamp(&mut beam, x, width, n);
        if params.dual {
            let other = match params.path {
                ScannerPath::Bounce => 1.0 - x,
                ScannerPath::Wrap => (x + 0.5).rem_euclid(1.0),
            };
            stamp(&mut beam, other, width, n);
        }
        for i in 0..n {
            trail[i] = trail[i].max(beam[i] * (0.35 + params.trail * 0.65));
        }

        let mut lamps_now = vec![[0u8; 3]; n];
        for i in 0..n {
            let t = lamps::pos(i, n);
            let e = (beam[i] + trail[i] * 0.75).max(params.field).clamp(0.0, 1.0);
            lamps_now[i] = lamps::scale_rgb(scanner_color(params.palette, rgb, t, hue_phase), e);
        }
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(33));
    }
}

fn stamp(out: &mut [f32], x: f32, width: f32, n: usize) {
    let center = x * n.saturating_sub(1) as f32;
    for (i, slot) in out.iter_mut().enumerate() {
        let dist = (i as f32 - center).abs();
        let amt = (-dist * dist / (2.0 * (width * n as f32).max(0.18).powi(2))).exp();
        *slot = (*slot + amt).min(1.0);
    }
}

fn scanner_color(palette: ScannerPalette, rgb: [u8; 12], t: f32, hue: f32) -> [u8; 3] {
    match palette {
        ScannerPalette::Custom => scene::custom(&rgb, t),
        ScannerPalette::Red => [255, 28, 18],
        ScannerPalette::Ice => scene::hsv(196.0 + t * 16.0, 0.7, 1.0),
        ScannerPalette::Rainbow => scene::hsv(hue + t * 80.0, 0.85, 1.0),
    }
}
