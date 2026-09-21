//! Column-native Audio React maps: Gravcenter, Melt, Wavelength.
//! Kept out of `audio.rs` so Classic 4-zone paint stays untouched.

use super::lamps;

pub fn bpm_from_ibi(ibi: f32, locked: bool) -> u16 {
    if !locked {
        0
    } else {
        (60.0 / ibi.max(0.2)).round().clamp(40.0, 240.0) as u16
    }
}

/// Instant rise, linear gravity fall. Smoothness 0 is a fast analog drop; 0.95 is slow.
pub fn tick_grav(height: &mut f32, peak: f32, dt: f32, smoothness: f32) {
    let peak = peak.clamp(0.0, 1.0);
    if peak > *height {
        *height = peak;
    } else {
        let gravity = (2.85 - smoothness.clamp(0.0, 0.95) * 2.5).max(0.28);
        *height = (*height - gravity * dt).max(peak * 0.1).max(0.0);
    }
}

/// Fill from the middle outward. `phase` adds a tiny sine wobble on the front.
pub fn grav_map(n: usize, height: f32, phase: f32) -> Vec<f32> {
    let n = n.max(1);
    let height = height.clamp(0.0, 1.0);
    let mut out = vec![0.0f32; n];
    let edge = (0.16 + 0.6 / n as f32).clamp(0.12, 0.28);
    for i in 0..n {
        let t = lamps::pos(i, n);
        let dist = (t - 0.5).abs() * 2.0;
        let wobble = 0.04 * (phase * 2.15 + i as f32 * 0.65).sin();
        let h = (height + wobble).clamp(0.0, 1.0);
        out[i] = ((h - dist) / edge).clamp(0.0, 1.0).powf(0.82);
    }
    out
}

pub fn tick_melt(heat: &mut Vec<f32>, inject: &[f32], dt: f32, smoothness: f32) {
    let n = inject.len();
    if heat.len() != n {
        *heat = vec![0.0; n];
    }
    if n == 0 {
        return;
    }
    for i in 0..n {
        heat[i] = (heat[i] + inject[i].clamp(0.0, 1.0) * dt * 4.4).min(1.0);
    }
    if n > 1 {
        let visc = 1.35 + smoothness.clamp(0.0, 0.95) * 2.4;
        let mut flow = vec![0.0f32; n - 1];
        for i in 0..n - 1 {
            flow[i] = (heat[i] - heat[i + 1]) * visc * dt;
        }
        for i in 0..n - 1 {
            heat[i] -= flow[i];
            heat[i + 1] += flow[i];
        }
    }
    let cool = (0.32 + (1.0 - smoothness.clamp(0.0, 0.95)) * 0.95) * dt;
    for slot in heat.iter_mut() {
        *slot = slot.clamp(0.0, 1.0);
        *slot = (*slot - cool).max(0.0);
    }
}

pub fn tick_wavelength(scroll: &mut f32, dt: f32, motion: f32) {
    *scroll = (*scroll + dt * (0.07 + motion.clamp(0.0, 2.0) * 0.24)).rem_euclid(1.0);
}

pub fn wavelength_energy(n: usize, bands: &[f32], scroll: f32) -> Vec<f32> {
    let n = n.max(1);
    let mut out = vec![0.0f32; n];
    if bands.is_empty() {
        return out;
    }
    let last = bands.len().saturating_sub(1) as f32;
    for i in 0..n {
        let t = (lamps::pos(i, n) + scroll.rem_euclid(1.0)).rem_euclid(1.0);
        out[i] = sample_span(bands, t * last);
    }
    out
}

pub fn wavelength_hue(t: f32, scroll: f32, centroid: f32) -> f32 {
    (t * 360.0 + scroll.rem_euclid(1.0) * 360.0 + centroid.clamp(0.0, 1.0) * 90.0).rem_euclid(360.0)
}

fn sample_span(bands: &[f32], src: f32) -> f32 {
    if bands.is_empty() {
        return 0.0;
    }
    let lo = src.floor() as usize;
    let hi = (lo + 1).min(bands.len() - 1);
    let f = (src - lo as f32).clamp(0.0, 1.0);
    bands[lo.min(bands.len() - 1)] * (1.0 - f) + bands[hi] * f
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn bpm_locked_half_second_ibi_is_120() {
        assert_eq!(bpm_from_ibi(0.5, true), 120);
        assert_eq!(bpm_from_ibi(0.5, false), 0);
    }

    #[test]
    fn grav_falls_after_the_peak() {
        let mut h = 0.0f32;
        tick_grav(&mut h, 0.9, 0.016, 0.4);
        assert!((h - 0.9).abs() < 1e-5);
        tick_grav(&mut h, 0.0, 0.05, 0.4);
        assert!(h < 0.9 && h > 0.0, "linear gravity should drop, not snap: {h}");
        let mid = h;
        tick_grav(&mut h, 0.0, 0.05, 0.4);
        assert!(h < mid, "should keep falling: {h} vs {mid}");
    }

    #[test]
    fn grav_center_brighter_than_edges() {
        let map = grav_map(24, 0.85, 0.0);
        assert!(map[11] > map[0] + 0.2);
        assert!(map[12] > map[23] + 0.2);
        let zone = grav_map(4, 0.8, 0.0);
        assert!(zone[1] > zone[0]);
        assert!(zone[2] > zone[3]);
    }

    #[test]
    fn melt_spreads_from_center() {
        let mut heat = vec![0.0f32; 8];
        let mut inject = vec![0.0f32; 8];
        inject[3] = 1.0;
        inject[4] = 1.0;
        for _ in 0..40 {
            tick_melt(&mut heat, &inject, 0.016, 0.5);
        }
        assert!(heat[3] > 0.15 && heat[4] > 0.15);
        assert!(heat[2] > 0.02 || heat[5] > 0.02, "viscous flow should leak sideways: {heat:?}");
        assert!(heat[0] < heat[3], "edges stay cooler than the inject: {heat:?}");
    }

    #[test]
    fn wavelength_scroll_moves_energy() {
        let bands: Vec<f32> = (0..24).map(|i| if i == 0 { 1.0 } else { 0.0 }).collect();
        let a = wavelength_energy(24, &bands, 0.0);
        let b = wavelength_energy(24, &bands, 0.5);
        let peak_a = a
            .iter()
            .enumerate()
            .max_by(|x, y| x.1.total_cmp(y.1))
            .map(|(i, _)| i)
            .unwrap();
        let peak_b = b
            .iter()
            .enumerate()
            .max_by(|x, y| x.1.total_cmp(y.1))
            .map(|(i, _)| i)
            .unwrap();
        assert_ne!(peak_a, peak_b, "scroll should move the bright column");
        let hue0 = wavelength_hue(0.0, 0.0, 0.2);
        let hue1 = wavelength_hue(0.0, 0.25, 0.2);
        assert!((hue1 - hue0).abs() > 40.0);
    }
}
