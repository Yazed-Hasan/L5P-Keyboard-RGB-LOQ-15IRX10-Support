use std::{sync::atomic::Ordering, thread, time::Duration};

use rand::Rng;

use crate::manager::Inner;

pub fn play(manager: &mut Inner, rng: &mut rand::rngs::ThreadRng) {
    let xmas_color_array = [[255, 10, 10], [255, 255, 20], [30, 255, 30], [70, 70, 255]];
    let subeffect_count = 4;
    let mut last_subeffect = -1;
    let n = manager.lamp_n();
    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        let mut subeffect = rng.random_range(0..subeffect_count);
        while last_subeffect == subeffect {
            subeffect = rng.random_range(0..subeffect_count);
        }
        last_subeffect = subeffect;

        match subeffect {
            0 => {
                for _i in 0..3 {
                    for colors in xmas_color_array {
                        manager.paint_lamps(&vec![colors; n]);
                        thread::sleep(Duration::from_millis(500));
                    }
                }
            }
            1 => {
                let color_1_index = rng.random_range(0..4);
                let used_colors_1: [u8; 3] = xmas_color_array[color_1_index];

                let mut color_2_index = rng.random_range(0..4);
                while color_1_index == color_2_index {
                    color_2_index = rng.random_range(0..4);
                }
                let used_colors_2: [u8; 3] = xmas_color_array[color_2_index];

                for _i in 0..4 {
                    manager.paint_lamps(&vec![used_colors_1; n]);
                    thread::sleep(Duration::from_millis(400));
                    manager.paint_lamps(&vec![used_colors_2; n]);
                    thread::sleep(Duration::from_millis(400));
                }
            }
            2 => {
                manager.paint_lamps(&vec![[0u8; 3]; n]);
                let range: Vec<usize> = if rng.random_range(0..2) == 0 {
                    (0..n).collect()
                } else {
                    (0..n).rev().collect()
                };
                for color in xmas_color_array {
                    let mut lamps = vec![[0u8; 3]; n];
                    for &j in &range {
                        lamps[j] = color;
                        manager.paint_lamps(&lamps);
                        thread::sleep(Duration::from_millis(18));
                        if manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
                            return;
                        }
                    }
                    for &j in &range {
                        lamps[j] = [0; 3];
                        manager.paint_lamps(&lamps);
                        thread::sleep(Duration::from_millis(18));
                        if manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
                            return;
                        }
                    }
                }
            }
            3 => {
                for _i in 0..4 {
                    let state1: Vec<[u8; 3]> = (0..n)
                        .map(|i| if i % 2 == 0 { [255, 255, 255] } else { [0, 0, 0] })
                        .collect();
                    let state2: Vec<[u8; 3]> = (0..n)
                        .map(|i| if i % 2 == 0 { [0, 0, 0] } else { [255, 255, 255] })
                        .collect();
                    manager.paint_lamps(&state1);
                    thread::sleep(Duration::from_millis(400));
                    manager.paint_lamps(&state2);
                    thread::sleep(Duration::from_millis(400));
                }
            }
            _ => unreachable!("Subeffect index for Christmas effect is out of range."),
        }
    }
}
