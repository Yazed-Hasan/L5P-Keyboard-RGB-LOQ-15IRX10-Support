use std::{
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use crate::{
    enums::{BatteryPalette, BatteryParams, Effects},
    manager::{effects::lamps, Inner},
};

use super::scene;

struct Power {
    charge: f32,
    ac: bool,
}

pub fn play(manager: &mut Inner, start: BatteryParams, rgb0: [u8; 12]) {
    let mut params = start.normalized();
    let mut rgb = rgb0;
    let mut shown = 1.0f32;
    let mut power = read_power();
    let mut last_poll = Instant::now();
    let mut last = Instant::now();
    let mut pulse_phase = 0.0f32;

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        if let Ok(g) = manager.scene_params.lock() {
            if let Effects::Battery { params: live } = g.effect {
                params = live.normalized();
                rgb = g.rgb;
            }
        }

        let dt = last.elapsed().as_secs_f32().clamp(0.008, 0.08);
        last = Instant::now();
        if last_poll.elapsed() >= Duration::from_millis(2000) {
            power = read_power();
            last_poll = Instant::now();
        }

        let follow = 1.0 - params.smoothing;
        shown += (power.charge - shown) * follow.clamp(0.04, 1.0);
        shown = shown.clamp(0.0, 1.0);
        pulse_phase = (pulse_phase + dt * (1.1 + params.pulse_speed * 2.4)).rem_euclid(std::f32::consts::TAU);

        let n = manager.lamp_n();
        let fill = shown;
        let tip = fill * n as f32;
        let pulse = if power.ac {
            (pulse_phase.cos() * 0.5 + 0.5).powf(1.6) * params.pulse
        } else {
            0.0
        };
        let color = battery_color(params.palette, rgb, fill, params.low_pct, params.mid_pct);

        let mut lamps_now = vec![[0u8; 3]; n];
        for i in 0..n {
            let idx = if params.reverse { n - 1 - i } else { i };
            let bar = (tip - i as f32).clamp(0.0, 1.0);
            let unused = if bar <= 0.001 { params.unused_dim } else { 0.0 };
            let near_tip = 1.0 - (i as f32 + 0.5 - tip).abs().clamp(0.0, 1.6) / 1.6;
            let e = (bar + unused + pulse * near_tip * bar.max(0.15)).clamp(0.0, 1.0);
            let shade = if bar > 0.001 {
                color
            } else {
                lamps::scale_rgb(color, 0.55)
            };
            lamps_now[idx] = lamps::scale_rgb(shade, e);
        }
        manager.paint_lamps(&lamps_now);
        thread::sleep(Duration::from_millis(33));
    }
}

fn battery_color(palette: BatteryPalette, rgb: [u8; 12], fill: f32, low: u8, mid: u8) -> [u8; 3] {
    match palette {
        BatteryPalette::Custom => scene::custom(&rgb, fill),
        BatteryPalette::Ice => scene::hsv(196.0 - fill * 20.0, 0.55, 1.0),
        BatteryPalette::Heat => scene::hsv((12.0 + fill * 48.0).min(55.0), 0.9, 1.0),
        BatteryPalette::Traffic => {
            let pct = fill * 100.0;
            if pct <= low as f32 {
                [255, 32, 24]
            } else if pct <= mid as f32 {
                [255, 188, 24]
            } else {
                [32, 220, 72]
            }
        }
    }
}

fn read_power() -> Power {
    #[cfg(windows)]
    {
        windows_power().unwrap_or(Power { charge: 1.0, ac: true })
    }
    #[cfg(not(windows))]
    {
        Power { charge: 1.0, ac: true }
    }
}

#[cfg(windows)]
fn windows_power() -> Option<Power> {
    use windows::Win32::System::Power::{GetSystemPowerStatus, SYSTEM_POWER_STATUS};

    let mut status = SYSTEM_POWER_STATUS::default();
    unsafe {
        GetSystemPowerStatus(&mut status).ok()?;
    }
    let no_battery = status.BatteryFlag == 128 || status.BatteryLifePercent > 100;
    if no_battery {
        return Some(Power { charge: 1.0, ac: true });
    }
    Some(Power {
        charge: status.BatteryLifePercent as f32 / 100.0,
        ac: status.ACLineStatus == 1,
    })
}
