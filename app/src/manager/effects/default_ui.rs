use eframe::egui::{ComboBox, Label, Sense, Slider, Ui};
use legion_rgb_driver::SPEED_RANGE;
use strum::IntoEnumIterator;

use crate::{
    enums::{Brightness, Direction},
    gui::style::SpacingStyle,
    manager::profile::Profile,
};

const COMBOBOX_WIDTH: f32 = 20.0;

fn tip(resp: eframe::egui::Response, text: &str) -> eframe::egui::Response {
    resp.on_hover_text(text).on_disabled_hover_text(text)
}

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
        let text = "Overall keyboard brightness for this effect.";
        ui.horizontal(|ui| {
            *update_lights |= tip(ui.add(Slider::new(&mut profile.brightness_level, 1..=100)), text).changed();
            tip(ui.add(Label::new("Brightness").sense(Sense::hover())), text);
            tip(ui.small_button("?"), text);
            if tip(
                ui.add_enabled(profile.brightness_level != 50, eframe::egui::Button::new("↺")),
                "Reset this setting to default",
            )
            .clicked()
            {
                profile.brightness_level = 50;
                *update_lights = true;
            }
        });
    } else {
        let text = "Firmware brightness steps. Software effects use the slider instead.";
        ui.horizontal(|ui| {
            tip(
                ComboBox::from_label("Brightness")
                    .width(COMBOBOX_WIDTH)
                    .selected_text({
                        let text: &'static str = profile.brightness.into();
                        text
                    })
                    .show_ui(ui, |ui| {
                        for val in Brightness::iter() {
                            let name: &'static str = val.into();
                            *update_lights |= ui.selectable_value(&mut profile.brightness, val, name).changed();
                        }
                    })
                    .response,
                text,
            );
            tip(ui.small_button("?"), text);
            if tip(
                ui.add_enabled(profile.brightness != Brightness::Low, eframe::egui::Button::new("↺")),
                "Reset this setting to default",
            )
            .clicked()
            {
                profile.brightness = Brightness::Low;
                *update_lights = true;
            }
        });
    }
}

pub fn show_direction(ui: &mut Ui, profile: &mut Profile, update_lights: &mut bool) {
    ui.add_enabled_ui(profile.effect.takes_direction(), |ui| {
        let text = "Which way Wave and Swipe travel across the four zones.";
        ui.horizontal(|ui| {
            tip(
                ComboBox::from_label("Direction")
                    .width(COMBOBOX_WIDTH)
                    .selected_text({
                        let text: &'static str = profile.direction.into();
                        text
                    })
                    .show_ui(ui, |ui| {
                        for val in Direction::iter() {
                            let name: &'static str = val.into();
                            *update_lights |= ui.selectable_value(&mut profile.direction, val, name).changed();
                        }
                    })
                    .response,
                text,
            );
            tip(ui.small_button("?"), text);
            if tip(
                ui.add_enabled(profile.direction != Direction::Left, eframe::egui::Button::new("↺")),
                "Reset this setting to default",
            )
            .clicked()
            {
                profile.direction = Direction::Left;
                *update_lights = true;
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

    let text = "How fast the effect animates. Higher is quicker. Applies live for software lighting.";
    ui.horizontal(|ui| {
        let changed = tip(
            ui.add_enabled(profile.effect.takes_speed(), Slider::new(&mut profile.speed, range)),
            text,
        )
        .changed();
        if changed {
            if is_dynamic_lighting {
                *live_speed = Some(profile.speed);
            } else {
                *update_lights = true;
            }
        }
        tip(ui.add(Label::new("Speed").sense(Sense::hover())), text);
        tip(ui.small_button("?"), text);
        if tip(
            ui.add_enabled(profile.effect.takes_speed() && profile.speed != 1, eframe::egui::Button::new("↺")),
            "Reset this setting to default",
        )
        .clicked()
        {
            profile.speed = 1;
            if is_dynamic_lighting {
                *live_speed = Some(1);
            } else {
                *update_lights = true;
            }
        }
    });
}
