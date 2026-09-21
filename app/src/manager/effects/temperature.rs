use std::{sync::atomic::Ordering, thread, time::Duration};

use sysinfo::{Components, System};

use crate::manager::{effects::lamps, Inner};

pub fn play(manager: &mut Inner) {
    let safe_temp = 20.0;
    let ramp_boost = 1.6;

    let mut sys = System::new_all();
    sys.refresh_all();

    let mut components = Components::new_with_refreshed_list();

    for component in components.iter_mut() {
        if component.label().contains("Tctl") {
            while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
                component.refresh();
                let temp = component.temperature();
                if let Some(temperature) = temp {
                    let mut adjusted_temp = temperature - safe_temp;
                    if adjusted_temp < 0.0 {
                        adjusted_temp = 0.0;
                    }
                    let temp_percent = ((adjusted_temp / 100.0) * ramp_boost).clamp(0.0, 1.0);
                    let n = manager.lamp_n();
                    let cool = [0u8, 255, 40];
                    let hot = [255u8, 20, 0];
                    let lamps_now: Vec<[u8; 3]> = (0..n)
                        .map(|i| {
                            let across = lamps::pos(i, n);
                            let t = (temp_percent * 0.75 + across * temp_percent * 0.45).clamp(0.0, 1.0);
                            lamps::mix_rgb(cool, hot, t)
                        })
                        .collect();
                    manager.paint_lamps(&lamps_now);
                }
                thread::sleep(Duration::from_millis(200));
            }
        }
    }
}
