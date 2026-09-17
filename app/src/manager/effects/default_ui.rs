use eframe::egui::{ComboBox, Slider, Ui};
use legion_rgb_driver::SPEED_RANGE;
use strum::IntoEnumIterator;

use crate::{
    enums::{Brightness, Direction},
    gui::style::SpacingStyle,
    manager::profile::Profile,
};

const COMBOBOX_WIDTH: f32 = 20.0;

pub fn show(
    ui: &mut Ui,
    profile: &mut Profile,
    update_lights: &mut bool,
    spacing: &SpacingStyle,
    is_dynamic_lighting: bool,
    live_speed: &mut Option<u8>,
) {
    ui.scope(|ui| {
        ui.style_mut().spacing.item_spacing = spacing.default;

        show_brightness(ui, profile, update_lights, is_dynamic_lighting);
        show_direction(ui, profile, update_lights);
        show_effect_settings(ui, profile, update_lights, is_dynamic_lighting, live_speed);
    });
}

pub fn show_brightness(ui: &mut Ui, profile: &mut Profile, update_lights: &mut bool, is_dynamic_lighting: bool) {
    if is_dynamic_lighting {
        ui.horizontal(|ui| {
            *update_lights |= ui.add(Slider::new(&mut profile.brightness_level, 1..=100)).changed();
            ui.label("Brightness");
        });
    } else {
        ComboBox::from_label("Brightness")
            .width(COMBOBOX_WIDTH)
            .selected_text({
                let text: &'static str = profile.brightness.into();
                text
            })
            .show_ui(ui, |ui| {
                for val in Brightness::iter() {
                    let text: &'static str = val.into();
                    *update_lights |= ui.selectable_value(&mut profile.brightness, val, text).changed();
                }
            });
    }
}

pub fn show_direction(ui: &mut Ui, profile: &mut Profile, update_lights: &mut bool) {
    ui.add_enabled_ui(profile.effect.takes_direction(), |ui| {
        ComboBox::from_label("Direction")
            .width(COMBOBOX_WIDTH)
            .selected_text({
                let text: &'static str = profile.direction.into();
                text
            })
            .show_ui(ui, |ui| {
                for val in Direction::iter() {
                    let text: &'static str = val.into();
                    *update_lights |= ui.selectable_value(&mut profile.direction, val, text).changed();
                }
            });
    });
}

pub fn show_effect_settings(
    ui: &mut Ui,
    profile: &mut Profile,
    update_lights: &mut bool,
    is_dynamic_lighting: bool,
    live_speed: &mut Option<u8>,
) {
    // Firmware effects only accept 1..=4. Software/WDL effects use 1..=10.
    let range = if is_dynamic_lighting || !profile.effect.is_built_in() {
        1..=10
    } else {
        SPEED_RANGE
    };

    ui.horizontal(|ui| {
        let changed = ui
            .add_enabled(profile.effect.takes_speed(), Slider::new(&mut profile.speed, range))
            .changed();
        if changed {
            if is_dynamic_lighting {
                *live_speed = Some(profile.speed);
            } else {
                *update_lights = true;
            }
        }
        ui.label("Speed");
    });
}
