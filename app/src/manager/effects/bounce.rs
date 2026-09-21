use std::{
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use rand::Rng;

use crate::{
    enums::{BouncePalette, BounceParams, Effects},
    manager::{effects::lamps, Inner},
};

use super::scene;

struct Ball {
    x: f32,
    v: f32,
    hue: f32,
}

pub fn play(manager: &mut Inner, start: BounceParams, rgb0: [u8; 12]) {
    legion_rgb_driver::debug_log("EFFECT: bouncing balls");
    let mut params = start.normalized();
    let mut rgb = rgb0;
    let mut trail = Vec::new();
    let mut balls: Vec<Ball> = Vec::new();
    let mut last = Instant::now();
    let mut rng = rand::rng();

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::Bounce { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        let n = manager.lamp_n();
        if trail.len() != n {
            trail = vec![0.0f32; n];
        }

        let want = params.count.round().clamp(1.0, 8.0) as usize;
        while balls.len() < want {
            balls.push(spawn_ball(&mut rng, balls.len()));
        }
        if balls.len() > want {
            balls.truncate(want);
        }

        let g = 1.15 + params.gravity * 2.6;
        for ball in balls.iter_mut() {
            step_ball(&mut ball.x, &mut ball.v, dt, g);
        }

        let decay = (1.5 + (1.0 - params.trail) * 4.8) * dt;
        for slot in trail.iter_mut() {
            *slot = (*slot - decay).max(0.0);
        }

        let mut beam = vec![0.0f32; n];
        let sigma = 0.35 + params.size * 0.9;
        for ball in &balls {
            stamp(&mut beam, ball.x, sigma, n);
        }
        for i in 0..n {
            trail[i] = trail[i].max(beam[i] * (0.28 + params.trail * 0.72));
        }

        let mut lamps_now = vec![[0u8; 3]; n];
        for i in 0..n {
            let t = lamps::pos(i, n);
            let e = (beam[i] + trail[i] * 0.78).max(params.background).clamp(0.0, 1.0);
            let hue = balls.first().map(|b| b.hue).unwrap_or(200.0);
            lamps_now[i] = lamps::scale_rgb(bounce_color(params.palette, rgb, t, hue + t * 40.0), e);
        }
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(if n > 4 { 33 } else { 24 }));
    }
}

pub fn step_ball(x: &mut f32, v: &mut f32, dt: f32, gravity: f32) {
    *v += gravity * dt;
    *x += *v * dt;
    if *x < 0.0 {
        *x = 0.0;
        *v = v.abs() * 0.86;
    } else if *x > 1.0 {
        *x = 1.0;
        *v = -v.abs() * 0.86;
    }
}

fn spawn_ball(rng: &mut impl Rng, i: usize) -> Ball {
    Ball {
        x: rng.random_range(0.08..0.92),
        v: if i % 2 == 0 { 0.35 } else { -0.2 },
        hue: rng.random_range(0.0..360.0),
    }
}

fn stamp(out: &mut [f32], x: f32, sigma: f32, n: usize) {
    if n == 0 {
        return;
    }
    let center = x.clamp(0.0, 1.0) * n.saturating_sub(1) as f32;
    let w = if n > 4 { sigma.max(0.28) } else { (sigma * 0.55).max(0.22) };
    for (i, slot) in out.iter_mut().enumerate() {
        let d = i as f32 - center;
        *slot = (*slot + (-d * d / (2.0 * w * w)).exp()).min(1.0);
    }
}

fn bounce_color(palette: BouncePalette, rgb: [u8; 12], t: f32, hue: f32) -> [u8; 3] {
    match palette {
        BouncePalette::Custom => scene::custom(&rgb, t),
        BouncePalette::Rainbow => scene::hsv(hue.rem_euclid(360.0), 0.88, 1.0),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn bounce_reverses_at_the_ends() {
        let mut x = 0.96;
        let mut v = 0.8;
        step_ball(&mut x, &mut v, 0.08, 1.2);
        assert!(x <= 1.0);
        assert!(v < 0.0, "should reverse at the right edge, v={v}");

        let mut x = 0.04;
        let mut v = -0.9;
        step_ball(&mut x, &mut v, 0.08, 1.2);
        assert!(x >= 0.0);
        assert!(v > 0.0, "should reverse at the left edge, v={v}");
    }
}
