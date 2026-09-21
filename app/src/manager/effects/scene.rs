use crate::enums::Effects;

use super::{audio_color, lamps};

#[derive(Clone, Copy)]
pub struct SceneLive {
    pub effect: Effects,
    pub rgb: [u8; 12],
}

impl Default for SceneLive {
    fn default() -> Self {
        Self {
            effect: Effects::Static,
            rgb: [255; 12],
        }
    }
}

pub fn hsv(h: f32, s: f32, v: f32) -> [u8; 3] {
    let (r, g, b) = audio_color::hsv_to_rgb(h.rem_euclid(360.0), s.clamp(0.0, 1.0), v.clamp(0.0, 1.0));
    [r, g, b]
}

pub fn custom(rgb: &[u8; 12], t: f32) -> [u8; 3] {
    lamps::sample_zones(rgb, t)
}

