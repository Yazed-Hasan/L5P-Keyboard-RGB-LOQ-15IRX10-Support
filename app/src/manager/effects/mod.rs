use default_ui::{show_brightness, show_direction, show_effect_settings};
use eframe::egui::{self, ComboBox, Slider};
use strum::IntoEnumIterator;

use crate::{
    enums::{AudioColorMode, AudioStyle, Effects, SwipeMode},
    manager::profile::Profile,
};

pub mod ambient;
pub mod audio;
pub mod christmas;
pub mod default_ui;
pub mod disco;
pub mod fade;
pub mod lightning;
pub mod ripple;
pub mod swipe;
pub mod temperature;
pub mod zones;

pub fn show_effect_ui(
    ui: &mut egui::Ui,
    profile: &mut Profile,
    update_lights: &mut bool,
    theme: &crate::gui::style::Theme,
    is_dynamic_lighting: bool,
    live_speed: &mut Option<u8>,
) {
    let mut effect = profile.effect;

    match &mut effect {
        Effects::SmoothWave { mode, clean_with_black } | Effects::Swipe { mode, clean_with_black } => {
            ui.scope(|ui| {
                ui.style_mut().spacing.item_spacing = theme.spacing.default;

                show_brightness(ui, profile, update_lights, is_dynamic_lighting);
                show_direction(ui, profile, update_lights);
                show_effect_settings(ui, profile, update_lights, is_dynamic_lighting, live_speed);
                ComboBox::from_label("Swipe mode").width(30.0).selected_text(format!("{:?}", mode)).show_ui(ui, |ui| {
                    for swipe_mode in SwipeMode::iter() {
                        *update_lights |= ui.selectable_value(mode, swipe_mode, format!("{:?}", swipe_mode)).changed();
                    }
                });
                *update_lights |= ui.add_enabled(matches!(mode, SwipeMode::Fill), egui::Checkbox::new(clean_with_black, "Clean with black")).changed();
            });
        }
        Effects::AudioReact {
            sensitivity,
            smoothness,
            min_brightness,
            bass,
            mid,
            treble,
            presence,
            squelch,
            punch,
            color_mode,
            style,
        } => {
            ui.scope(|ui| {
                ui.style_mut().spacing.item_spacing = theme.spacing.default;
                show_brightness(ui, profile, update_lights, is_dynamic_lighting);

                *sensitivity = if *sensitivity <= 0.05 { 1.2 } else { *sensitivity };
                *smoothness = smoothness.clamp(0.0, 0.95);
                if *bass <= 0.001 {
                    *bass = 1.0;
                }
                if *mid <= 0.001 {
                    *mid = 1.0;
                }
                if *treble <= 0.001 {
                    *treble = 1.0;
                }
                if *presence <= 0.001 {
                    *presence = 1.0;
                }
                if *squelch < 0.0 {
                    *squelch = 0.07;
                }
                if *punch < 0.0 {
                    *punch = 0.65;
                }

                ui.horizontal(|ui| {
                    ui.add(Slider::new(sensitivity, 0.2..=5.0));
                    ui.label("Sensitivity");
                });
                ui.horizontal(|ui| {
                    ui.add(Slider::new(smoothness, 0.0..=0.95));
                    ui.label("Smoothness");
                });
                ui.horizontal(|ui| {
                    ui.add(Slider::new(squelch, 0.0..=0.35));
                    ui.label("Squelch");
                });
                ui.horizontal(|ui| {
                    ui.add(Slider::new(punch, 0.0..=2.0));
                    ui.label("Punch");
                });
                ui.horizontal(|ui| {
                    ui.add(Slider::new(min_brightness, 0..=80));
                    ui.label("Idle glow");
                });
                ui.horizontal(|ui| {
                    ui.add(Slider::new(bass, 0.0..=2.5));
                    ui.label("Bass");
                });
                ui.horizontal(|ui| {
                    ui.add(Slider::new(mid, 0.0..=2.5));
                    ui.label("Mid");
                });
                ui.horizontal(|ui| {
                    ui.add(Slider::new(treble, 0.0..=2.5));
                    ui.label("Treble");
                });
                ui.horizontal(|ui| {
                    ui.add(Slider::new(presence, 0.0..=2.5));
                    ui.label("Presence");
                });
                ComboBox::from_label("Color mode")
                    .width(80.0)
                    .selected_text(format!("{:?}", color_mode))
                    .show_ui(ui, |ui| {
                        for mode in AudioColorMode::iter() {
                            ui.selectable_value(color_mode, mode, format!("{:?}", mode));
                        }
                    });
                ComboBox::from_label("Style")
                    .width(80.0)
                    .selected_text(format!("{:?}", style))
                    .show_ui(ui, |ui| {
                        for value in AudioStyle::iter() {
                            ui.selectable_value(style, value, format!("{:?}", value));
                        }
                    });
            });
        }
        Effects::AmbientLight { fps, saturation_boost } => {
            ui.scope(|ui| {
                ui.style_mut().spacing.item_spacing = theme.spacing.default;

                show_brightness(ui, profile, update_lights, is_dynamic_lighting);
                show_direction(ui, profile, update_lights);

                ui.horizontal(|ui| {
                    *update_lights |= ui.add(Slider::new(fps, 1..=60)).changed();
                    ui.label("FPS");
                });
                ui.horizontal(|ui| {
                    *update_lights |= ui.add(Slider::new(saturation_boost, 0.0..=1.0)).changed();
                    ui.label("Saturation Boost");
                });
            });
        }
        _ => {
            default_ui::show(ui, profile, update_lights, &theme.spacing, is_dynamic_lighting, live_speed);
        }
    }

    profile.effect = effect;
}
