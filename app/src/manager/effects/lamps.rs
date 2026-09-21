use legion_rgb_driver::MAX_LAMPS;

pub fn count(lamp_count: u16) -> usize {
    (lamp_count as usize).clamp(1, MAX_LAMPS)
}

pub fn pos(i: usize, n: usize) -> f32 {
    if n <= 1 {
        0.0
    } else {
        i as f32 / (n - 1) as f32
    }
}

pub fn lerp_u8(a: u8, b: u8, t: f32) -> u8 {
    (a as f32 + (b as f32 - a as f32) * t.clamp(0.0, 1.0)).round() as u8
}

pub fn mix_rgb(a: [u8; 3], b: [u8; 3], t: f32) -> [u8; 3] {
    [lerp_u8(a[0], b[0], t), lerp_u8(a[1], b[1], t), lerp_u8(a[2], b[2], t)]
}

pub fn scale_rgb(c: [u8; 3], amt: f32) -> [u8; 3] {
    let a = amt.clamp(0.0, 1.0);
    [
        (c[0] as f32 * a) as u8,
        (c[1] as f32 * a) as u8,
        (c[2] as f32 * a) as u8,
    ]
}

pub fn zone_rgb(rgb: &[u8; 12], z: usize) -> [u8; 3] {
    let z = z.min(3);
    [rgb[z * 3], rgb[z * 3 + 1], rgb[z * 3 + 2]]
}

pub fn sample_zones(rgb: &[u8; 12], t: f32) -> [u8; 3] {
    let x = t.clamp(0.0, 1.0) * 3.0;
    let i = x.floor() as usize;
    let f = x - i as f32;
    mix_rgb(zone_rgb(rgb, i), zone_rgb(rgb, i + 1), f)
}

pub fn sample_bands(levels: &[f32; 4], t: f32) -> f32 {
    let x = t.clamp(0.0, 1.0) * 3.0;
    let i = x.floor() as usize;
    let f = x - i as f32;
    let a = levels[i.min(3)];
    let b = levels[(i + 1).min(3)];
    a * (1.0 - f) + b * f
}

pub fn lamps_from_zones(rgb: &[u8; 12], n: usize) -> Vec<[u8; 3]> {
    (0..n).map(|i| sample_zones(rgb, pos(i, n))).collect()
}
