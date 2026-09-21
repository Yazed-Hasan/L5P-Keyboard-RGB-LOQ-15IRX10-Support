use std::{
    f32::consts::PI,
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use crate::{
    enums::{CometPalette, CometParams, Direction, Effects, ScannerPath},
    manager::{effects::lamps, Inner},
};

use super::scene;

pub fn play(manager: &mut Inner, start: CometParams, rgb0: [u8; 12]) {
    legion_rgb_driver::debug_log("EFFECT: comet");
    let mut params = start.normalized();
    let mut rgb = rgb0;
    let going_left = matches!(start.direction, Direction::Left);
    let mut x = if going_left { 1.0 } else { 0.0 };
    let mut dir = if going_left { -1.0 } else { 1.0 };
    let mut travel = if going_left { 1.0 } else { 0.0 };
    let mut trail: Vec<f32> = Vec::new();
    let mut last = Instant::now();
    let mut hue_phase = 18.0f32;

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::Comet { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        hue_phase = (hue_phase + dt * (18.0 + params.speed * 36.0) * params.hue_speed).rem_euclid(360.0);
        let n = manager.lamp_n().max(1);
        if trail.len() != n {
            trail = vec![0.0f32; n];
        }

        let wrap = matches!(params.path, ScannerPath::Wrap);
        let wobble = 1.0 - params.wobble * 0.12 + params.wobble * 0.28 * (travel * 2.4f32).sin().abs();
        let speed = (0.22 + params.speed * 0.95) * wobble;

        let mut tail_amt = 1.0f32;
        if wrap {
            dir = if matches!(params.direction, Direction::Left) {
                -1.0
            } else {
                1.0
            };
            travel += dir * speed * dt;
            x = travel.rem_euclid(1.0);
        } else {
            travel += speed * dt;
            let bounced = ping_pong(travel);
            x = bounced.x;
            dir = bounced.dir;
            tail_amt = bounced.tail_amt;
        }

        let fade_rate = (1.15 + (1.0 - params.tail) * 3.6) * dt;
        for slot in trail.iter_mut() {
            *slot = (*slot - fade_rate).max(0.0);
        }

        let mut energy = paint_comet(x, dir, n, &params, wrap, hue_phase, tail_amt);
        if params.dual {
            let extra = extra_body(travel, dir, &params, wrap);
            let second = paint_comet(extra.x, extra.dir, n, &params, wrap, hue_phase + 40.0, extra.tail_amt);
            for (slot, add) in energy.iter_mut().zip(second) {
                *slot = (*slot).max(add * params.follow);
            }
            if params.triple {
                let mut third_params = params;
                third_params.gap = (params.gap * 1.85).min(0.84);
                let extra = extra_body(travel, dir, &third_params, wrap);
                let third = paint_comet(extra.x, extra.dir, n, &params, wrap, hue_phase + 80.0, extra.tail_amt);
                for (slot, add) in energy.iter_mut().zip(third) {
                    *slot = (*slot).max(add * (params.follow * 0.72));
                }
            }
        }

        for (slot, e) in trail.iter_mut().zip(energy.iter()) {
            *slot = (*slot).max(*e * (0.42 + params.tail * 0.4));
        }

        let lamps_now: Vec<[u8; 3]> = (0..n)
            .map(|i| {
                let t = lamps::pos(i, n);
                let e = energy[i].max(trail[i]).max(params.background).clamp(0.0, 1.0);
                lamps::scale_rgb(comet_color(params.palette, rgb, t, hue_phase, e, params.saturation), e)
            })
            .collect();
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(if n > 4 { 24 } else { 16 }));
    }
}

struct Body {
    x: f32,
    dir: f32,
    tail_amt: f32,
}

fn ping_pong(travel: f32) -> Body {
    let cycle = travel.rem_euclid(2.0);
    let (u, dir) = if cycle <= 1.0 {
        (cycle, 1.0)
    } else {
        (2.0 - cycle, -1.0)
    };
    let x = 0.5 - 0.5 * (u * PI).cos();
    let speed = (u * PI).sin().abs();
    Body { x, dir, tail_amt: 0.18 + 0.82 * speed }
}

fn extra_body(travel: f32, dir: f32, params: &CometParams, wrap: bool) -> Body {
    if wrap {
        let d = if params.opposite { -dir } else { dir };
        let x = if params.opposite {
            (travel + dir * params.gap).rem_euclid(1.0)
        } else {
            (travel - dir * params.gap).rem_euclid(1.0)
        };
        Body { x, dir: d, tail_amt: 1.0 }
    } else {
        let delay = if params.opposite { 1.0 } else { params.gap };
        ping_pong(travel - delay)
    }
}

fn paint_comet(
    x: f32,
    dir: f32,
    n: usize,
    params: &CometParams,
    wrap: bool,
    phase: f32,
    tail_amt: f32,
) -> Vec<f32> {
    let mut out = vec![0.0f32; n];
    if n == 0 {
        return out;
    }
    let span = n as f32;
    let center = x.clamp(0.0, 1.0) * (span - 1.0).max(0.0);
    let sigma = (0.22 + params.size * 1.25).max(0.18);
    let glow_sigma = sigma * (1.8 + params.glow * 2.4);
    let tail_len = (1.15 + params.tail * span * 1.05).max(1.0);
    let dir = if dir < 0.0 { -1.0 } else { 1.0 };
    let fade = params.fade.max(0.45);
    let tail_amt = tail_amt.clamp(0.0, 1.0);
    for i in 0..n {
        let pos = i as f32;
        let delta = pos - center;
        let head = (-delta * delta / (2.0 * sigma * sigma)).exp() * params.head;
        let halo = if params.glow > 0.01 {
            (-delta * delta / (2.0 * glow_sigma * glow_sigma)).exp() * params.glow * 0.55
        } else {
            0.0
        };
        let mut behind = (center - pos) * dir;
        if wrap {
            behind = behind.rem_euclid(span);
        }
        let streak = if behind > 0.1 && behind < tail_len {
            let t = (1.0 - behind / tail_len).clamp(0.0, 1.0);
            t.powf(fade) * (0.28 + params.tail * 0.72) * tail_amt
        } else {
            0.0
        };
        let spark = if params.sparkle > 0.01 && streak > 0.04 {
            let nse = ((i as f32 * 17.3 + phase * 0.11).sin() * 0.5 + 0.5).powf(4.0);
            nse * params.sparkle * streak
        } else {
            0.0
        };
        out[i] = (head + halo + streak + spark).min(1.0);
    }
    out
}

fn comet_color(palette: CometPalette, rgb: [u8; 12], t: f32, hue: f32, energy: f32, saturation: f32) -> [u8; 3] {
    let sat = saturation.clamp(0.15, 1.0);
    let raw = match palette {
        CometPalette::Custom => scene::custom(&rgb, t),
        CometPalette::Heat => {
            let hot = scene::hsv(18.0 + energy * 28.0 + hue * 0.04, 0.95, 1.0);
            let core = scene::hsv(42.0, 0.35, 1.0);
            lamps::mix_rgb(hot, core, energy.powf(2.2) * 0.55)
        }
        CometPalette::Ice => scene::hsv(196.0 + energy * 18.0 + hue * 0.05, 0.55, 1.0),
        CometPalette::Rainbow => scene::hsv(hue + t * 90.0, 0.88, 1.0),
    };
    if sat >= 0.995 {
        return raw;
    }
    let grey = ((raw[0] as u16 + raw[1] as u16 + raw[2] as u16) / 3) as u8;
    lamps::mix_rgb(raw, [grey, grey, grey], 1.0 - sat)
}
