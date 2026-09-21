use std::{sync::atomic::Ordering, thread, time::Duration};

use rand::{rngs::ThreadRng, Rng};

use crate::manager::{effects::lamps, profile::Profile, Inner};

pub fn play(manager: &mut Inner, p: &Profile, rng: &mut ThreadRng) {
    let n = manager.lamp_n();
    let base = if p.rgb_array().iter().any(|v| *v >= 24) {
        p.rgb_array()
    } else {
        [255, 220, 80, 255, 180, 40, 255, 255, 200, 180, 220, 255]
    };
    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
            break;
        }
        let center = rng.random_range(0..n);
        let color = lamps::sample_zones(&base, lamps::pos(center, n));
        let width = (n as f32 / 5.5).max(0.8);
        let mut lamps_now = vec![[0u8; 3]; n];
        for i in 0..n {
            let dist = (i as f32 - center as f32).abs();
            let amt = (-dist * dist / (2.0 * width * width)).exp();
            lamps_now[i] = lamps::scale_rgb(color, amt);
        }
        manager.paint_lamps(&lamps_now);
        let steps = (8 + rng.random_range(4..12)).max(1);
        for step in 1..=steps {
            if manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
                return;
            }
            let fade = 1.0 - step as f32 / steps as f32;
            let faded: Vec<[u8; 3]> = lamps_now.iter().copied().map(|c| lamps::scale_rgb(c, fade)).collect();
            manager.paint_lamps(&faded);
            thread::sleep(Duration::from_millis(18));
        }
        manager.paint_lamps(&vec![[0u8; 3]; n]);
        let sleep_time = rng.random_range(80..=700) / p.speed.max(1) as u64 * 4;
        thread::sleep(Duration::from_millis(sleep_time.max(40)));
    }
}
