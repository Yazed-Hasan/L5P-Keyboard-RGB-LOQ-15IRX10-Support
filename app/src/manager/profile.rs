use std::{convert::TryInto, path::Path};

use crate::{
    enums::{Brightness, Direction, Effects},
    util::StorageTrait,
};

use error_stack::{Result, ResultExt};
use legion_rgb_driver::{expand_zones_to_lamps, pack_lamp_rgb, unpack_lamp_rgb, zone_colors_from_lamps, LAMP_ZONES, MAX_LAMPS};
use serde::{Deserialize, Serialize};
use thiserror::Error;

#[derive(Clone, Copy, Serialize, Deserialize, Debug, PartialEq, Eq)]
pub struct KeyboardZone {
    pub rgb: [u8; 3],
    pub enabled: bool,
}

impl Default for KeyboardZone {
    fn default() -> Self {
        Self {
            rgb: Default::default(),
            enabled: true,
        }
    }
}

type Zones = [KeyboardZone; 4];

#[derive(Serialize, Deserialize, Clone, Debug, PartialEq)]
pub struct Profile {
    pub name: Option<String>,
    pub rgb_zones: Zones,
    pub effect: Effects,
    pub direction: Direction,
    pub speed: u8,
    pub brightness: Brightness,
    #[serde(default = "default_brightness_level")]
    pub brightness_level: u8,
    /// Per-lamp RGB (3 bytes per lamp). Empty means expand from the 4 zones.
    #[serde(default)]
    pub lamp_rgb: Vec<u8>,
}

fn default_brightness_level() -> u8 {
    50
}

impl Default for Profile {
    fn default() -> Self {
        Self {
            name: None,
            rgb_zones: Zones::default(),
            effect: Effects::default(),
            direction: Direction::default(),
            speed: 1,
            brightness: Brightness::default(),
            brightness_level: 50,
            lamp_rgb: Vec::new(),
        }
    }
}

#[derive(Debug, Error)]
#[error("Could not load profile")]
pub struct LoadProfileError;

#[derive(Debug, Error)]
#[error("Could not save profile")]
pub struct SaveProfileError;

impl Profile {
    pub fn load_profile(path: &Path) -> Result<Self, LoadProfileError> {
        Self::load(path).change_context(LoadProfileError)
    }

    pub fn save_profile(&mut self, path: &Path) -> Result<(), SaveProfileError> {
        if self.name.is_none() {
            self.name = Some("Untitled".to_string());
        }
        crate::persist::export_profile(self, path)
            .map(|_| ())
            .map_err(|_| error_stack::Report::new(SaveProfileError))
    }

    pub fn rgb_array(&self) -> [u8; 12] {
        self.rgb_zones.map(|zone| if zone.enabled { zone.rgb } else { [0; 3] }).concat().try_into().unwrap()
    }

    pub fn has_lamp_colors(&self) -> bool {
        self.lamp_rgb.len() >= 3
    }

    pub fn lamp_colors(&self, lamp_count: usize) -> Vec<[u8; 3]> {
        let lc = lamp_count.max(1).min(MAX_LAMPS);
        if self.lamp_rgb.len() >= 3 {
            unpack_lamp_rgb(&self.lamp_rgb, lc)
        } else {
            expand_zones_to_lamps(&self.rgb_array(), lc)
        }
    }

    pub fn ensure_lamps(&mut self, lamp_count: usize) {
        let lc = lamp_count.max(1).min(MAX_LAMPS);
        if self.lamp_rgb.len() != lc * 3 {
            let packed = pack_lamp_rgb(&expand_zones_to_lamps(&self.rgb_array(), lc));
            self.lamp_rgb = packed[..lc * 3].to_vec();
        }
    }

    pub fn set_lamp(&mut self, index: usize, rgb: [u8; 3], lamp_count: usize) {
        self.ensure_lamps(lamp_count);
        let lc = lamp_count.max(1).min(MAX_LAMPS);
        if index >= lc {
            return;
        }
        let o = index * 3;
        if o + 2 < self.lamp_rgb.len() {
            self.lamp_rgb[o] = rgb[0];
            self.lamp_rgb[o + 1] = rgb[1];
            self.lamp_rgb[o + 2] = rgb[2];
        }
        self.sync_zones_from_lamps(lc);
    }

    pub fn fill_zone_lamps(&mut self, zone: usize, lamp_count: usize) {
        self.ensure_lamps(lamp_count);
        let lc = lamp_count.max(1).min(MAX_LAMPS);
        let zones = LAMP_ZONES.min(lc);
        if zone >= zones {
            return;
        }
        let color = if self.rgb_zones[zone].enabled {
            self.rgb_zones[zone].rgb
        } else {
            [0, 0, 0]
        };
        let start = (zone * lc) / zones;
        let end = ((zone + 1) * lc) / zones;
        for i in start..end {
            let o = i * 3;
            if o + 2 < self.lamp_rgb.len() {
                self.lamp_rgb[o] = color[0];
                self.lamp_rgb[o + 1] = color[1];
                self.lamp_rgb[o + 2] = color[2];
            }
        }
    }

    pub fn fill_all_lamps(&mut self, rgb: [u8; 3], lamp_count: usize) {
        self.ensure_lamps(lamp_count);
        for chunk in self.lamp_rgb.chunks_mut(3) {
            if chunk.len() == 3 {
                chunk.copy_from_slice(&rgb);
            }
        }
        for zone in &mut self.rgb_zones {
            zone.rgb = rgb;
        }
    }

    pub fn clear_lamp_colors(&mut self) {
        self.lamp_rgb.clear();
    }

    fn sync_zones_from_lamps(&mut self, lamp_count: usize) {
        let lamps = unpack_lamp_rgb(&self.lamp_rgb, lamp_count);
        let rgb = zone_colors_from_lamps(&lamps);
        self.rgb_zones = arr_to_zones(rgb);
    }
}

pub fn arr_to_zones(arr: [u8; 12]) -> Zones {
    [
        KeyboardZone {
            rgb: arr[0..3].try_into().unwrap(),
            enabled: true,
        },
        KeyboardZone {
            rgb: arr[3..6].try_into().unwrap(),
            enabled: true,
        },
        KeyboardZone {
            rgb: arr[6..9].try_into().unwrap(),
            enabled: true,
        },
        KeyboardZone {
            rgb: arr[9..12].try_into().unwrap(),
            enabled: true,
        },
    ]
}

pub const DEFAULT_AUDIO_ZONE_RGB: [u8; 12] = [255, 24, 48, 255, 140, 16, 36, 220, 120, 72, 120, 255];

impl Profile {
    pub fn reset_current_mode(&mut self) {
        self.effect = self.effect.factory_default();
        if self.effect.takes_speed() {
            self.speed = 1;
        }
        if self.effect.takes_direction() {
            self.direction = Direction::Left;
        }
        if matches!(self.effect, Effects::AudioReact { .. }) {
            self.rgb_zones = arr_to_zones(DEFAULT_AUDIO_ZONE_RGB);
            self.lamp_rgb.clear();
        }
    }
}

impl StorageTrait<'_> for Profile {}
