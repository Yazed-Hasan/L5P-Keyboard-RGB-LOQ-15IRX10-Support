use std::{
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use crate::{
    enums::{AuroraPalette, AuroraParams, Effects},
    manager::{effects::lamps, Inner},
};

use super::scene;

pub fn play(manager: &mut Inner, start: AuroraParams, rgb0: [u8; 12]) {
    let mut params = start.normalized();
    let mut rgb = rgb0;
    let mut last = Instant::now();
    let mut time = 0.0f32;

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::Aurora { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        time += dt * params.speed;
        let n = manager.lamp_n();
        let layers = params.layers as usize;
        let wave = 1.6 + params.wavelength * 4.2;
        let mut raw = vec![0.0f32; n];
        let mut hue_mix = vec![0.0f32; n];

        for i in 0..n {
            let t = lamps::pos(i, n);
            let mut e = 0.0;
            let mut h = 0.0;
            for layer in 0..layers {
                let phase = time * (0.7 + layer as f32 * 0.45) + t * wave * (0.55 + layer as f32 * 0.22);
                let s = (phase.sin() * 0.5 + 0.5).powf(1.15);
                e += s / layers as f32;
                h += (140.0 + layer as f32 * 38.0 + t * 50.0) / layers as f32;
            }
            let shaped = ((e - 0.5) * params.contrast + 0.5).clamp(0.0, 1.0);
            raw[i] = shaped * params.brightness;
            hue_mix[i] = h + time * params.hue_drift * 36.0;
        }

        if params.softness > 0.01 {
            let mut blur = raw.clone();
            let k = 0.15 + params.softness * 0.45;
            for i in 0..n {
                let mut acc = raw[i] * (1.0 - k);
                if i > 0 {
                    acc += raw[i - 1] * k * 0.5;
                }
                if i + 1 < n {
                    acc += raw[i + 1] * k * 0.5;
                }
                blur[i] = acc;
            }
            raw = blur;
        }

        let mut lamps_now = vec![[0u8; 3]; n];
        for i in 0..n {
            let t = lamps::pos(i, n);
            lamps_now[i] = lamps::scale_rgb(aurora_color(params.palette, rgb, t, hue_mix[i], raw[i]), raw[i]);
        }
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(33));
    }
}

fn aurora_color(palette: AuroraPalette, rgb: [u8; 12], t: f32, hue: f32, energy: f32) -> [u8; 3] {
    match palette {
        AuroraPalette::Custom => scene::custom(&rgb, t),
        AuroraPalette::Borealis => {
            let h = 130.0 + t * 70.0 + energy * 24.0;
            scene::hsv(h, 0.72, 1.0)
        }
        AuroraPalette::Twilight => scene::hsv(268.0 + t * 50.0 - energy * 18.0, 0.62, 1.0),
        AuroraPalette::Rainbow => scene::hsv(hue, 0.7, 1.0),
    }
}
