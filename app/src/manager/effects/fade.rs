use std::{
    sync::{
        atomic::{AtomicBool, Ordering},
        Arc,
    },
    thread,
    time::{Duration, Instant},
};

use device_query::DeviceQuery;

use crate::manager::{effects::lamps, profile::Profile, Inner};

pub fn play(manager: &mut Inner, p: &Profile) {
    let stop_signals = manager.stop_signals.clone();

    let kill_thread = Arc::new(AtomicBool::new(false));
    let exit_thread = kill_thread.clone();

    let state = device_query::DeviceState::new();

    thread::spawn(move || {
        let state = device_query::DeviceState::new();

        loop {
            if !state.get_keys().is_empty() {
                stop_signals.keyboard_stop_signal.store(true, Ordering::SeqCst);
            }

            if exit_thread.load(Ordering::SeqCst) {
                break;
            }

            thread::sleep(Duration::from_millis(5));
        }
    });

    let mut now = Instant::now();
    let n = manager.lamp_n();
    let base = p.lamp_colors(n);
    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if state.get_keys().is_empty() {
            if now.elapsed() > Duration::from_secs(20 / u64::from(p.speed.max(1))) {
                let steps = 40u8;
                for step in 1..=steps {
                    if manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
                        break;
                    }
                    let amt = 1.0 - step as f32 / steps as f32;
                    let faded: Vec<[u8; 3]> = base.iter().copied().map(|c| lamps::scale_rgb(c, amt)).collect();
                    manager.paint_lamps(&faded);
                    thread::sleep(Duration::from_millis(18));
                }
                manager.paint_lamps(&vec![[0u8; 3]; n]);
            } else {
                thread::sleep(Duration::from_millis(20));
            }
        } else {
            manager.paint_lamps(&base);
            manager.stop_signals.keyboard_stop_signal.store(false, Ordering::SeqCst);
            now = Instant::now();
        }
    }

    kill_thread.store(true, Ordering::SeqCst);
}
