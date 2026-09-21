use std::{sync::atomic::Ordering, thread, time::Duration};

use rand::Rng;

use crate::manager::{effects::lamps, profile::Profile, Inner};

pub fn play(manager: &mut Inner, p: &Profile, rng: &mut rand::rngs::ThreadRng) {
    let n = manager.lamp_n();
    let palette = [[255, 0, 0], [255, 255, 0], [0, 255, 0], [0, 255, 255], [0, 0, 255], [255, 0, 255]];
    let mut current = vec![[0u8; 3]; n];
    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        for lamp in current.iter_mut() {
            *lamp = lamps::scale_rgb(*lamp, 0.7);
        }
        current[rng.random_range(0..n)] = palette[rng.random_range(0..palette.len())];
        if n > 4 {
            current[rng.random_range(0..n)] = palette[rng.random_range(0..palette.len())];
        }
        manager.paint_lamps(&current);
        thread::sleep(Duration::from_millis((2000 / (u64::from(p.speed.max(1)) * 4)).max(40)));
    }
}
