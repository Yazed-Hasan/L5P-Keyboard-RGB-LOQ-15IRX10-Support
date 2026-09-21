use std::{
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use rand::Rng;

use crate::{
    enums::{DissolvePalette, DissolveParams, Effects},
    manager::{effects::lamps, Inner},
};

use super::scene;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Phase {
    Fill,
    Hold,
    Clear,
}

pub fn play(manager: &mut Inner, start: DissolveParams, rgb0: [u8; 12]) {
    legion_rgb_driver::debug_log("EFFECT: dissolve");
    let mut params = start.normalized();
    let mut rgb = rgb0;
    let mut last = Instant::now();
    let mut rng = rand::rng();
    let mut order: Vec<usize> = Vec::new();
    let mut colors: Vec<[u8; 3]> = Vec::new();
    let mut lit = Vec::new();
    let mut cursor = 0usize;
    let mut phase = Phase::Fill;
    let mut hold = 0.0f32;
    let mut hue_phase = 0.0f32;

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::Dissolve { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        hue_phase = (hue_phase + dt * 28.0).rem_euclid(360.0);
        let n = manager.lamp_n();
        if order.len() != n {
            order = shuffled(n, &mut rng);
            colors = (0..n)
                .map(|i| dissolve_color(params.palette, params.random_colors, rgb, lamps::pos(i, n), hue_phase, &mut rng))
                .collect();
            lit = vec![false; n];
            cursor = 0;
            phase = Phase::Fill;
            hold = 0.0;
        }

        let rate = 4.0 + params.speed * 14.0;
        match phase {
            Phase::Fill => {
                cursor = advance_fill(&mut lit, &order, cursor, dt, rate);
                if filled_count(&lit) >= n {
                    phase = Phase::Hold;
                    hold = 0.0;
                }
            }
            Phase::Hold => {
                hold += dt;
                if hold > 0.28 {
                    order = shuffled(n, &mut rng);
                    cursor = 0;
                    phase = Phase::Clear;
                }
            }
            Phase::Clear => {
                cursor = advance_clear(&mut lit, &order, cursor, dt, rate);
                if filled_count(&lit) == 0 {
                    order = shuffled(n, &mut rng);
                    colors = (0..n)
                        .map(|i| dissolve_color(params.palette, params.random_colors, rgb, lamps::pos(i, n), hue_phase, &mut rng))
                        .collect();
                    cursor = 0;
                    phase = Phase::Fill;
                }
            }
        }

        let mut lamps_now = vec![[0u8; 3]; n];
        for i in 0..n {
            let t = lamps::pos(i, n);
            let base = dissolve_color(params.palette, false, rgb, t, hue_phase, &mut rng);
            let bg = lamps::scale_rgb(base, params.background);
            lamps_now[i] = if lit[i] { colors[i] } else { bg };
        }
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(if n > 4 { 33 } else { 24 }));
    }
}

pub fn shuffled(n: usize, rng: &mut impl Rng) -> Vec<usize> {
    let mut order: Vec<usize> = (0..n).collect();
    for i in (1..order.len()).rev() {
        let j = rng.random_range(0..=i);
        order.swap(i, j);
    }
    order
}

pub fn advance_fill(lit: &mut [bool], order: &[usize], cursor: usize, dt: f32, rate: f32) -> usize {
    step_flags(lit, order, cursor, dt, rate, true)
}

pub fn advance_clear(lit: &mut [bool], order: &[usize], cursor: usize, dt: f32, rate: f32) -> usize {
    step_flags(lit, order, cursor, dt, rate, false)
}

pub fn filled_count(lit: &[bool]) -> usize {
    lit.iter().filter(|v| **v).count()
}

fn step_flags(lit: &mut [bool], order: &[usize], mut cursor: usize, dt: f32, rate: f32, on: bool) -> usize {
    let n = lit.len();
    if n == 0 || order.len() != n {
        return cursor;
    }
    let mut steps = (rate * dt).floor() as usize;
    if rng_frac(rate * dt) {
        steps += 1;
    }
    steps = steps.max(1);
    for _ in 0..steps {
        if cursor >= n {
            break;
        }
        let i = order[cursor];
        if i < n {
            lit[i] = on;
        }
        cursor += 1;
    }
    cursor
}

fn rng_frac(_frac: f32) -> bool {
    false
}

fn dissolve_color(
    palette: DissolvePalette,
    random_colors: bool,
    rgb: [u8; 12],
    t: f32,
    hue: f32,
    rng: &mut impl Rng,
) -> [u8; 3] {
    if random_colors {
        return scene::hsv(rng.random_range(0.0..360.0), 0.9, 1.0);
    }
    match palette {
        DissolvePalette::Custom => scene::custom(&rgb, t),
        DissolvePalette::Rainbow => scene::hsv((hue + t * 120.0).rem_euclid(360.0), 0.9, 1.0),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rand::SeedableRng;

    #[test]
    fn dissolve_eventually_fills_all() {
        let n = 24;
        let mut rng = rand::rngs::StdRng::seed_from_u64(3);
        let order = shuffled(n, &mut rng);
        assert_eq!(order.len(), n);
        let mut seen = vec![false; n];
        for &i in &order {
            seen[i] = true;
        }
        assert!(seen.iter().all(|&v| v));

        let mut lit = vec![false; n];
        let mut cursor = 0usize;
        for _ in 0..40 {
            cursor = advance_fill(&mut lit, &order, cursor, 0.05, 20.0);
        }
        assert_eq!(filled_count(&lit), n);
        cursor = 0;
        for _ in 0..40 {
            cursor = advance_clear(&mut lit, &order, cursor, 0.05, 20.0);
        }
        assert_eq!(filled_count(&lit), 0);
    }
}
