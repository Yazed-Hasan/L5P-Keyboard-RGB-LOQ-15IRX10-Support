use std::{
    f32::consts::TAU,
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use crate::{
    enums::{Effects, JugglePalette, JuggleParams},
    manager::{effects::lamps, Inner},
};

use super::scene;

pub fn play(manager: &mut Inner, start: JuggleParams, rgb0: [u8; 12]) {
    legion_rgb_driver::debug_log("EFFECT: juggle");
    let mut params = start.normalized();
    let mut rgb = rgb0;
    let mut trail = Vec::new();
    let mut last = Instant::now();
    let mut time = 0.0f32;

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::Juggle { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        time = (time + dt).rem_euclid(4096.0);
        let n = manager.lamp_n();
        if trail.len() != n {
            trail = vec![0.0f32; n];
        }

        let decay = (1.4 + (1.0 - params.trail) * 5.2) * dt;
        for slot in trail.iter_mut() {
            *slot = (*slot - decay).max(0.0);
        }

        let dots = params.dots.round().clamp(2.0, 8.0) as usize;
        let mut beam = vec![0.0f32; n];
        for d in 0..dots {
            let x = juggle_x(d, time, params.speed);
            stamp(&mut beam, x, n);
        }
        for i in 0..n {
            trail[i] = trail[i].max(beam[i] * (0.3 + params.trail * 0.7));
        }

        let mut lamps_now = vec![[0u8; 3]; n];
        for i in 0..n {
            let t = lamps::pos(i, n);
            let e = (beam[i] + trail[i] * 0.8).max(params.background).clamp(0.0, 1.0);
            lamps_now[i] = lamps::scale_rgb(juggle_color(params.palette, rgb, t, time, i), e);
        }
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(if n > 4 { 33 } else { 24 }));
    }
}

pub fn juggle_x(dot: usize, time: f32, speed: f32) -> f32 {
    let freq = 0.32 + dot as f32 * 0.11 + speed.clamp(0.15, 2.4) * 0.55;
    (0.5 + 0.5 * (time * freq * TAU + dot as f32 * 1.7).sin()).clamp(0.0, 1.0)
}

fn stamp(out: &mut [f32], x: f32, n: usize) {
    if n == 0 {
        return;
    }
    let center = x.clamp(0.0, 1.0) * n.saturating_sub(1) as f32;
    let w = if n > 4 { 0.85f32 } else { 0.45 };
    for (i, slot) in out.iter_mut().enumerate() {
        let d = i as f32 - center;
        *slot = (*slot + (-d * d / (2.0 * w * w)).exp()).min(1.0);
    }
}

fn juggle_color(palette: JugglePalette, rgb: [u8; 12], t: f32, time: f32, i: usize) -> [u8; 3] {
    match palette {
        JugglePalette::Custom => scene::custom(&rgb, t),
        JugglePalette::Rainbow => scene::hsv((time * 40.0 + i as f32 * 18.0).rem_euclid(360.0), 0.88, 1.0),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn juggle_stays_on_the_strip() {
        for d in 0..8 {
            for step in 0..40 {
                let x = juggle_x(d, step as f32 * 0.05, 0.7);
                assert!((0.0..=1.0).contains(&x), "dot {d} t={step} x={x}");
            }
        }
    }
}
