use default_ui::{show_brightness, show_direction, show_effect_settings};
use eframe::egui::{self, ComboBox, Slider};
use strum::IntoEnumIterator;

use crate::{
    enums::{
        AudioColorMode, AudioStyle, AuroraPalette, BatteryPalette, Direction, Effects, RainPalette, RippleKind, RippleOrigin,
        RippleTint, RippleTrigger, ScannerPalette, ScannerPath, StarsPalette, SwipeMode,
    },
    manager::profile::Profile,
};

pub mod ambient;
pub mod audio;
pub mod audio_color;
pub mod aurora;
pub mod battery;
pub mod christmas;
pub mod default_ui;
pub mod disco;
pub mod fade;
pub mod lamps;
pub mod lightning;
pub mod rain;
pub mod ripple;
pub(crate) mod scene;
pub mod scanner;
pub mod stars;
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
                combo_tip(
                    ui,
                    ComboBox::from_label("Swipe mode").width(30.0).selected_text(format!("{:?}", mode)),
                    "Change swaps colors as the wave moves. Fill paints the keyboard then clears it.",
                    |ui| {
                        for swipe_mode in SwipeMode::iter() {
                            *update_lights |= apply_tip(
                                ui.selectable_value(mode, swipe_mode, format!("{:?}", swipe_mode)),
                                swipe_mode_tip(swipe_mode),
                            )
                            .changed();
                        }
                    },
                );
                *update_lights |= apply_tip(
                    ui.add_enabled(matches!(mode, SwipeMode::Fill), egui::Checkbox::new(clean_with_black, "Clean with black")),
                    "After a Fill swipe, fade through black instead of the next color.",
                )
                .changed();
            });
        }
        Effects::AudioReact {
            sensitivity,
            smoothness,
            min_brightness,
            idle_brightness,
            bass,
            mid,
            treble,
            presence,
            squelch,
            punch,
            contrast,
            spread,
            hue_shift,
            color_ramp,
            motion,
            follow_system_volume,
            ripple_color,
            ripple_strength,
            ripple_speed,
            ripple_width,
            ripple_twist,
            ripple_origin,
            ripple_tint,
            ripple_rgb,
            ripple_kind,
            ripple_trigger,
            ripple_shockwave,
            ripple_shock_strength,
            ripple_shock_sensitivity,
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
                if *contrast <= 0.05 {
                    *contrast = 1.15;
                }
                if *spread < 0.0 {
                    *spread = 0.35;
                }
                if *hue_shift < 0.0 {
                    *hue_shift = 0.25;
                }
                if *color_ramp < 0.0 {
                    *color_ramp = 0.4;
                }
                if *motion < 0.0 {
                    *motion = 0.7;
                }

                if matches!(*style, AudioStyle::Ripple) {
                    *style = AudioStyle::Levels;
                    *ripple_color = true;
                }

                slider_tip(ui, sensitivity, 0.2..=5.0, "Sensitivity", "How hard the lights jump when the music gets louder. Raise this if quiet songs barely move.", 1.2);
                slider_tip(ui, smoothness, 0.0..=0.95, "Smoothness", "How slowly the lights fade after a beat. Lower is snappier; higher is softer and a bit later.", 0.72);
                slider_tip(ui, contrast, 0.5..=2.2, "Contrast", "Makes quiet parts dimmer and loud parts brighter. Higher looks more dramatic.", 1.15);
                slider_tip(ui, spread, 0.0..=1.0, "Spread", "Blurs energy into neighboring zones so the keyboard feels like one wash instead of four blocks.", 0.35);
                slider_tip(ui, motion, 0.0..=2.0, "Motion", "How fast traveling styles (Wave, Chase, Sparkle) and Ripple color rings move across the keys.", 0.7);
                slider_tip(ui, hue_shift, 0.0..=1.0, "Hue shift", "How much the color tint changes with the music. 0 keeps your colors; higher cycles them.", 0.25);
                slider_tip(ui, color_ramp, 0.0..=1.0, "Color ramp", "How slowly colors blend from one tint to the next. 0 snaps; higher eases a smooth ramp. Use with Ramp color mode for a drifting wash.", 0.4);
                slider_tip(ui, squelch, 0.0..=0.35, "Squelch", "Ignores faint hiss and background noise. Raise it if lights twitch when nothing is playing.", 0.07);
                slider_tip(ui, punch, 0.0..=2.0, "Punch", "Extra flash on sudden beats and drops. Higher hits harder on kicks.", 0.65);
                slider_tip(ui, min_brightness, 0..=80, "Beat floor", "Minimum light mixed into beats while music is playing. This is not used after the audio stops.", 6);
                {
                    if idle_brightness.is_none() {
                        *idle_brightness = Some(*min_brightness);
                    }
                    let idle = idle_brightness.get_or_insert(*min_brightness);
                    slider_tip(
                        ui,
                        idle,
                        0..=80,
                        "Idle glow",
                        "Light when audio has stopped. Independent from Beat floor — 0 is fully dark, higher keeps a rest glow.",
                        6,
                    );
                }
                slider_tip(ui, bass, 0.0..=2.5, "Bass", "Kick and bass (about 20–160 Hz). Turn up for more left-zone thump.", 1.0);
                slider_tip(ui, mid, 0.0..=2.5, "Mid", "Vocals and instruments (about 160–600 Hz).", 1.0);
                slider_tip(ui, treble, 0.0..=2.5, "Treble", "Hats, cymbals, and brightness (about 600–2500 Hz).", 1.0);
                slider_tip(ui, presence, 0.0..=2.5, "Presence", "Air and sparkle at the top end (about 2.5–9 kHz).", 1.0);
                combo_tip(
                    ui,
                    ComboBox::from_label("Color mode").width(100.0).selected_text(audio_color::name(*color_mode)),
                    "Where colors come from: your swatches, a palette, a moving rainbow, or loudness.",
                    |ui| {
                        for mode in AudioColorMode::iter() {
                            apply_tip(
                                ui.selectable_value(color_mode, mode, audio_color::name(mode)),
                                audio_color::tip(mode),
                            );
                        }
                    },
                );
                if let Some((label, hint)) = audio_color::swatch_hint(*color_mode) {
                    apply_tip(ui.label(label), hint);
                }
                combo_tip(
                    ui,
                    ComboBox::from_label("Style").width(90.0).selected_text(audio_style_name(*style)),
                    "How the lights are arranged. Hover each style in the list, or the (?) next to this row.",
                    |ui| {
                        for value in AudioStyle::iter().filter(|value| !matches!(value, AudioStyle::Ripple)) {
                            apply_tip(
                                ui.selectable_value(style, value, audio_style_name(value)),
                                audio_style_tip(value),
                            );
                        }
                    },
                );
                ui.horizontal(|ui| {
                    const TIP: &str = "On: a color ring spreads from each beat on top of the current Style. Mix with Wave, Levels, Chase, and the rest. Extra settings appear when this is on.";
                    apply_tip(ui.checkbox(ripple_color, "Ripple color"), TIP);
                    apply_tip(ui.small_button("?"), TIP);
                    if reset_btn(ui, *ripple_color, false) {
                        *ripple_color = false;
                    }
                });
                if *ripple_color {
                    ui.indent("ripple_color_settings", |ui| {
                        if *ripple_strength <= 0.0 {
                            *ripple_strength = 1.0;
                        }
                        if *ripple_speed <= 0.05 {
                            *ripple_speed = 1.0;
                        }
                        if *ripple_width <= 0.05 {
                            *ripple_width = 0.45;
                        }
                        if *ripple_twist < 0.0 {
                            *ripple_twist = 0.7;
                        }
                        slider_tip(ui, ripple_strength, 0.0..=2.0, "Ripple strength", "How bright the ring is on top of the current Style.", 1.0);
                        slider_tip(ui, ripple_speed, 0.2..=2.5, "Ripple speed", "How fast the ring travels across the keyboard. Independent from Motion.", 1.0);
                        slider_tip(ui, ripple_width, 0.1..=1.0, "Ripple width", "How thick the ring is. In 24-lamp mode this is strip count: low is one column, higher is a few columns. 4-zone stays a wider wash.", 0.45);
                        combo_tip(
                            ui,
                            ComboBox::from_label("Ripple type").width(110.0).selected_text(ripple_kind_name(*ripple_kind)),
                            "Shape of the ripple: a single ring, a traveling wave, a pulse, two rings, a fill, or echoes.",
                            |ui| {
                                for value in RippleKind::iter() {
                                    apply_tip(
                                        ui.selectable_value(ripple_kind, value, ripple_kind_name(value)),
                                        ripple_kind_tip(value),
                                    );
                                }
                            },
                        );
                        combo_tip(
                            ui,
                            ComboBox::from_label("Ripple color").width(110.0).selected_text(ripple_tint_name(*ripple_tint)),
                            "What color the ripple uses: keep the Style colors, shift hue, pick a color, or a rainbow.",
                            |ui| {
                                for value in RippleTint::iter() {
                                    apply_tip(
                                        ui.selectable_value(ripple_tint, value, ripple_tint_name(value)),
                                        ripple_tint_tip(value),
                                    );
                                }
                            },
                        );
                        if matches!(*ripple_tint, RippleTint::Custom) {
                            ui.horizontal(|ui| {
                                const TIP: &str = "The exact color of the ripple ring. Mixes over the current Style where the ring is.";
                                apply_tip(ui.color_edit_button_srgb(ripple_rgb), TIP);
                                setting_label(ui, "Ripple swatch", TIP);
                                if reset_btn(ui, *ripple_rgb, [255, 48, 96]) {
                                    *ripple_rgb = [255, 48, 96];
                                }
                            });
                        }
                        if matches!(*ripple_tint, RippleTint::ColorChange | RippleTint::Rainbow) {
                            slider_tip(ui, ripple_twist, 0.0..=1.5, "Color twist", "How much the ring's hue shifts as it expands. 0 keeps the starting color.", 0.7);
                        }
                        combo_tip(
                            ui,
                            ComboBox::from_label("Ripple on").width(110.0).selected_text(ripple_trigger_name(*ripple_trigger)),
                            "When a ring starts: any beat, bass, kick drum, or the third kick in a short burst.",
                            |ui| {
                                for value in RippleTrigger::iter() {
                                    apply_tip(
                                        ui.selectable_value(ripple_trigger, value, ripple_trigger_name(value)),
                                        ripple_trigger_tip(value),
                                    );
                                }
                            },
                        );
                        combo_tip(
                            ui,
                            ComboBox::from_label("Ripple origin").width(90.0).selected_text(ripple_origin_name(*ripple_origin)),
                            "Where each ring starts: loudest band, center, left, or right.",
                            |ui| {
                                for value in RippleOrigin::iter() {
                                    apply_tip(
                                        ui.selectable_value(ripple_origin, value, ripple_origin_name(value)),
                                        ripple_origin_tip(value),
                                    );
                                }
                            },
                        );
                        ui.horizontal(|ui| {
                            const TIP: &str = "On: a wide expanding blast fires on bass drops — the sudden kick after a quieter dip — on top of normal beat rings.";
                            apply_tip(ui.checkbox(ripple_shockwave, "Shockwave on drop"), TIP);
                            apply_tip(ui.small_button("?"), TIP);
                            if reset_btn(ui, *ripple_shockwave, false) {
                                *ripple_shockwave = false;
                            }
                        });
                        if *ripple_shockwave {
                            ui.indent("ripple_shock_settings", |ui| {
                                if *ripple_shock_strength <= 0.05 {
                                    *ripple_shock_strength = 1.25;
                                }
                                if *ripple_shock_sensitivity < 0.0 {
                                    *ripple_shock_sensitivity = 0.55;
                                }
                                slider_tip(
                                    ui,
                                    ripple_shock_strength,
                                    0.3..=2.0,
                                    "Shockwave strength",
                                    "How bright and wide the drop blast is. Higher fills more of the keyboard behind the expanding front.",
                                    1.25,
                                );
                                slider_tip(
                                    ui,
                                    ripple_shock_sensitivity,
                                    0.0..=1.0,
                                    "Drop sensitivity",
                                    "How easily a bass hit counts as a drop. Lower needs a clearer quiet-then-kick; higher fires more often.",
                                    0.55,
                                );
                            });
                        }
                    });
                }
                ui.horizontal(|ui| {
                    const TIP: &str = "Off (default): keep reacting to the music even at low Windows volume. On: lights stay full strength until Windows volume is about 5% or muted, then they stop.";
                    apply_tip(ui.checkbox(follow_system_volume, "Follow system volume"), TIP);
                    apply_tip(ui.small_button("?"), TIP);
                    if reset_btn(ui, *follow_system_volume, false) {
                        *follow_system_volume = false;
                    }
                });
            });
        }
        Effects::Stars { params } => {
            ui.scope(|ui| {
                ui.style_mut().spacing.item_spacing = theme.spacing.default;
                show_brightness(ui, profile, update_lights, is_dynamic_lighting);
                slider_tip(ui, &mut params.density, 0.05..=1.0, "Density", "How many stars are lit at once. Lower is a sparse night sky.", 0.45);
                slider_tip(ui, &mut params.twinkle, 0.2..=2.0, "Twinkle", "How fast each star fades in and out.", 0.7);
                slider_tip(ui, &mut params.size, 0.05..=1.0, "Size", "How wide each star spills onto neighboring strips. More visible in 24-lamp mode.", 0.35);
                slider_tip(ui, &mut params.background, 0.0..=0.4, "Background", "Dim wash behind the stars so the keyboard is not fully dark.", 0.06);
                slider_tip(ui, &mut params.shooting, 0.0..=0.3, "Shooting stars", "Chance a streak crosses the keys. 0 turns them off.", 0.08);
                slider_tip(ui, &mut params.hue_drift, 0.0..=1.0, "Hue drift", "Slowly shifts Rainbow and Random star colors.", 0.15);
                combo_tip(
                    ui,
                    ComboBox::from_label("Palette").width(90.0).selected_text(stars_palette_name(params.palette)),
                    "Custom uses the four zone swatches. Other palettes ignore them.",
                    |ui| {
                        for pal in StarsPalette::iter() {
                            apply_tip(ui.selectable_value(&mut params.palette, pal, stars_palette_name(pal)), stars_palette_tip(pal));
                        }
                    },
                );
            });
        }
        Effects::Rain { params } => {
            ui.scope(|ui| {
                ui.style_mut().spacing.item_spacing = theme.spacing.default;
                show_brightness(ui, profile, update_lights, is_dynamic_lighting);
                slider_tip(ui, &mut params.density, 0.05..=1.0, "Density", "How many raindrops are on the keyboard at once.", 0.55);
                slider_tip(ui, &mut params.speed, 0.15..=2.5, "Fall speed", "How fast drops travel across the keys.", 0.85);
                slider_tip(ui, &mut params.trail, 0.05..=1.0, "Trail", "Length of the streak behind each drop.", 0.45);
                slider_tip(ui, &mut params.splash, 0.0..=1.0, "Splash", "Flash at the far edge when a drop lands.", 0.7);
                slider_tip(ui, &mut params.wind, 0.0..=1.0, "Wind", "Adds speed jitter so drops do not fall in lockstep.", 0.15);
                slider_tip(ui, &mut params.wet, 0.0..=1.0, "Wet leftover", "How long faint rain stays after a drop passes.", 0.2);
                combo_tip(
                    ui,
                    ComboBox::from_label("Direction").width(70.0).selected_text(format!("{:?}", params.direction)),
                    "Which way the rain travels across the zones.",
                    |ui| {
                        for dir in Direction::iter() {
                            ui.selectable_value(&mut params.direction, dir, format!("{dir:?}"));
                        }
                    },
                );
                combo_tip(
                    ui,
                    ComboBox::from_label("Palette").width(90.0).selected_text(rain_palette_name(params.palette)),
                    "Custom uses the four zone swatches. Ice and Neon ignore them.",
                    |ui| {
                        for pal in RainPalette::iter() {
                            apply_tip(ui.selectable_value(&mut params.palette, pal, rain_palette_name(pal)), rain_palette_tip(pal));
                        }
                    },
                );
            });
        }
        Effects::Aurora { params } => {
            ui.scope(|ui| {
                ui.style_mut().spacing.item_spacing = theme.spacing.default;
                show_brightness(ui, profile, update_lights, is_dynamic_lighting);
                slider_tip(ui, &mut params.layers, 1..=4, "Layers", "How many overlapping aurora bands are mixed together.", 3);
                slider_tip(ui, &mut params.speed, 0.05..=2.0, "Speed", "How fast the bands drift.", 0.45);
                slider_tip(ui, &mut params.wavelength, 0.15..=2.0, "Wavelength", "How wide each band is across the keyboard.", 0.7);
                slider_tip(ui, &mut params.contrast, 0.4..=2.2, "Contrast", "Separates bright ridges from dark gaps.", 1.1);
                slider_tip(ui, &mut params.hue_drift, 0.0..=1.0, "Hue drift", "How quickly Rainbow and mixed hues wander.", 0.25);
                slider_tip(ui, &mut params.softness, 0.0..=1.0, "Softness", "Blurs neighboring strips so bands look washed, not blocky.", 0.45);
                slider_tip(ui, &mut params.brightness, 0.15..=1.0, "Glow", "Overall strength of the aurora on top of keyboard brightness.", 0.85);
                combo_tip(
                    ui,
                    ComboBox::from_label("Palette").width(90.0).selected_text(aurora_palette_name(params.palette)),
                    "Borealis is green-cyan. Twilight is purple. Custom uses your zone colors.",
                    |ui| {
                        for pal in AuroraPalette::iter() {
                            apply_tip(ui.selectable_value(&mut params.palette, pal, aurora_palette_name(pal)), aurora_palette_tip(pal));
                        }
                    },
                );
            });
        }
        Effects::Scanner { params } => {
            ui.scope(|ui| {
                ui.style_mut().spacing.item_spacing = theme.spacing.default;
                show_brightness(ui, profile, update_lights, is_dynamic_lighting);
                slider_tip(ui, &mut params.width, 0.06..=0.8, "Width", "How wide the moving hotspot is.", 0.28);
                slider_tip(ui, &mut params.speed, 0.1..=2.5, "Speed", "How fast the beam travels.", 0.7);
                slider_tip(ui, &mut params.trail, 0.0..=1.0, "Trail", "Fade left behind the beam.", 0.55);
                slider_tip(ui, &mut params.field, 0.0..=0.4, "Field", "Dim glow on the rest of the keyboard.", 0.08);
                ui.horizontal(|ui| {
                    const TIP: &str = "Adds a second beam coming the other way (bounce) or opposite side (wrap).";
                    apply_tip(ui.checkbox(&mut params.dual, "Dual beam"), TIP);
                    apply_tip(ui.small_button("?"), TIP);
                    if reset_btn(ui, params.dual, false) {
                        params.dual = false;
                    }
                });
                combo_tip(
                    ui,
                    ComboBox::from_label("Path").width(80.0).selected_text(scanner_path_name(params.path)),
                    "Bounce turns around at the edges. Wrap leaves one side and re-enters the other.",
                    |ui| {
                        for path in ScannerPath::iter() {
                            apply_tip(ui.selectable_value(&mut params.path, path, scanner_path_name(path)), scanner_path_tip(path));
                        }
                    },
                );
                combo_tip(
                    ui,
                    ComboBox::from_label("Direction").width(70.0).selected_text(format!("{:?}", params.direction)),
                    "Starting travel direction. Bounce still reverses at the ends.",
                    |ui| {
                        for dir in Direction::iter() {
                            ui.selectable_value(&mut params.direction, dir, format!("{dir:?}"));
                        }
                    },
                );
                combo_tip(
                    ui,
                    ComboBox::from_label("Palette").width(90.0).selected_text(scanner_palette_name(params.palette)),
                    "Red is the classic scanner. Custom uses the zone swatches.",
                    |ui| {
                        for pal in ScannerPalette::iter() {
                            apply_tip(ui.selectable_value(&mut params.palette, pal, scanner_palette_name(pal)), scanner_palette_tip(pal));
                        }
                    },
                );
            });
        }
        Effects::Battery { params } => {
            ui.scope(|ui| {
                ui.style_mut().spacing.item_spacing = theme.spacing.default;
                show_brightness(ui, profile, update_lights, is_dynamic_lighting);
                slider_tip(ui, &mut params.low_pct, 5..=40, "Low %", "Charge at or below this is the red/low color in Traffic palette.", 15);
                slider_tip(ui, &mut params.mid_pct, 20..=90, "Mid %", "Charge at or below this is the yellow/mid color in Traffic palette.", 50);
                slider_tip(ui, &mut params.unused_dim, 0.0..=0.35, "Unused dim", "How bright the empty part of the meter stays.", 0.06);
                slider_tip(ui, &mut params.pulse, 0.0..=1.5, "Charge pulse", "How hard the tip flashes while the laptop is on AC power.", 0.7);
                slider_tip(ui, &mut params.pulse_speed, 0.15..=2.5, "Pulse speed", "How fast the charging pulse breathes.", 0.8);
                slider_tip(ui, &mut params.smoothing, 0.0..=0.95, "Smoothing", "How slowly the bar follows charge changes. Higher is calmer.", 0.45);
                ui.horizontal(|ui| {
                    const TIP: &str = "Fill from the right instead of the left.";
                    apply_tip(ui.checkbox(&mut params.reverse, "Reverse fill"), TIP);
                    apply_tip(ui.small_button("?"), TIP);
                    if reset_btn(ui, params.reverse, false) {
                        params.reverse = false;
                    }
                });
                combo_tip(
                    ui,
                    ComboBox::from_label("Palette").width(90.0).selected_text(battery_palette_name(params.palette)),
                    "Traffic is green/yellow/red by charge. Custom uses the zone swatches along the fill.",
                    |ui| {
                        for pal in BatteryPalette::iter() {
                            apply_tip(ui.selectable_value(&mut params.palette, pal, battery_palette_name(pal)), battery_palette_tip(pal));
                        }
                    },
                );
            });
        }
        Effects::AmbientLight { fps, saturation_boost } => {
            ui.scope(|ui| {
                ui.style_mut().spacing.item_spacing = theme.spacing.default;

                show_brightness(ui, profile, update_lights, is_dynamic_lighting);
                show_direction(ui, profile, update_lights);

                ui.horizontal(|ui| {
                    const TIP: &str = "How often the keyboard samples the screen. Higher is smoother and uses more CPU.";
                    *update_lights |= apply_tip(ui.add(Slider::new(fps, 1..=60)), TIP).changed();
                    setting_label(ui, "FPS", TIP);
                    if reset_btn(ui, *fps, 24) {
                        *fps = 24;
                        *update_lights = true;
                    }
                });
                ui.horizontal(|ui| {
                    const TIP: &str = "Pushes sampled screen colors toward richer RGB so the keys look less washed out.";
                    *update_lights |= apply_tip(ui.add(Slider::new(saturation_boost, 0.0..=1.0)), TIP).changed();
                    setting_label(ui, "Saturation Boost", TIP);
                    if reset_btn(ui, *saturation_boost, 0.2) {
                        *saturation_boost = 0.2;
                        *update_lights = true;
                    }
                });
            });
        }
        _ => {
            default_ui::show(ui, profile, update_lights, &theme.spacing, is_dynamic_lighting, live_speed);
        }
    }

    profile.effect = effect;

    ui.horizontal(|ui| {
        const TIP: &str = "Restore this lighting mode's factory settings. Other modes keep their own saved values.";
        if apply_tip(ui.button("Reset this mode"), TIP).clicked() {
            profile.reset_current_mode();
            *update_lights = true;
        }
        apply_tip(ui.small_button("?"), TIP);
    });
}

fn apply_tip(response: egui::Response, tip: &str) -> egui::Response {
    response.on_hover_text(tip).on_disabled_hover_text(tip)
}

fn reset_btn<T: PartialEq>(ui: &mut egui::Ui, current: T, default: T) -> bool {
    apply_tip(
        ui.add_enabled(current != default, egui::Button::new("↺")),
        "Reset this setting to default",
    )
    .clicked()
}

fn setting_label(ui: &mut egui::Ui, label: &str, tip: &str) {
    apply_tip(
        ui.add(egui::Label::new(label).sense(egui::Sense::hover())),
        tip,
    );
    apply_tip(ui.small_button("?"), tip);
}

fn slider_tip<Num: egui::emath::Numeric>(
    ui: &mut egui::Ui,
    value: &mut Num,
    range: std::ops::RangeInclusive<Num>,
    label: &str,
    tip: &str,
    default: Num,
) {
    ui.horizontal(|ui| {
        let is_default = *value == default;
        apply_tip(ui.add(Slider::new(value, range)), tip);
        setting_label(ui, label, tip);
        if apply_tip(
            ui.add_enabled(!is_default, egui::Button::new("↺")),
            "Reset this setting to default",
        )
        .clicked()
        {
            *value = default;
        }
    });
}

fn combo_tip(ui: &mut egui::Ui, combo: ComboBox, tip: &str, add_contents: impl FnOnce(&mut egui::Ui)) {
    ui.horizontal(|ui| {
        apply_tip(combo.show_ui(ui, add_contents).response, tip);
        apply_tip(ui.small_button("?"), tip);
    });
}

fn audio_style_name(style: AudioStyle) -> &'static str {
    match style {
        AudioStyle::Levels => "Levels",
        AudioStyle::Pulse => "Pulse",
        AudioStyle::Wave => "Wave",
        AudioStyle::Bloom => "Bloom",
        AudioStyle::Center => "Center",
        AudioStyle::Mirror => "Mirror",
        AudioStyle::Fire => "Fire",
        AudioStyle::Strobe => "Strobe",
        AudioStyle::Sparkle => "Sparkle",
        AudioStyle::Chase => "Chase",
        AudioStyle::Gradient => "Gradient",
        AudioStyle::BeatGates => "Beat Gates",
        AudioStyle::Vu => "VU Meter",
        AudioStyle::TempoPulse => "Tempo Pulse",
        AudioStyle::Ripple => "Ripple",
    }
}

fn audio_style_tip(style: AudioStyle) -> &'static str {
    match style {
        AudioStyle::Levels => "Each zone or strip follows its own frequency band (bass on the left, presence on the right).",
        AudioStyle::Pulse => "The whole keyboard flashes together with overall loudness.",
        AudioStyle::Wave => "A traveling wave whose height follows the beat.",
        AudioStyle::Bloom => "Bass lights the left and spills right as it gets louder.",
        AudioStyle::Center => "Energy builds from the middle zones outward.",
        AudioStyle::Mirror => "Left and right zones mirror each other.",
        AudioStyle::Fire => "Bass at the left, hotter colors/energy climbing toward the right.",
        AudioStyle::Strobe => "Sharp flashes on peaks. Best with Punch turned up.",
        AudioStyle::Sparkle => "Random keys glitter on beats.",
        AudioStyle::Chase => "A moving hotspot that races with the music.",
        AudioStyle::Gradient => "A color/brightness ramp that slides with overall energy.",
        AudioStyle::BeatGates => "Each zone or strip snaps fully on or off with its own beat.",
        AudioStyle::Vu => "A studio meter that fills from the left with loudness. The bright tip is the recent peak.",
        AudioStyle::TempoPulse => "The whole keyboard pulses on the detected downbeat. Locks to BPM; falls back to loudness until a beat is found.",
        AudioStyle::Ripple => "A color ring spreads from the beat and shifts hue as it expands.",
    }
}

fn ripple_origin_name(origin: RippleOrigin) -> &'static str {
    match origin {
        RippleOrigin::Auto => "Auto",
        RippleOrigin::Center => "Center",
        RippleOrigin::Left => "Left",
        RippleOrigin::Right => "Right",
    }
}

fn ripple_origin_tip(origin: RippleOrigin) -> &'static str {
    match origin {
        RippleOrigin::Auto => "Start at the loudest band or strip. Kick and Triple kick start on the left.",
        RippleOrigin::Center => "Always start in the middle and spread both ways.",
        RippleOrigin::Left => "Always start on the left and travel right.",
        RippleOrigin::Right => "Always start on the right and travel left.",
    }
}

fn ripple_tint_name(tint: RippleTint) -> &'static str {
    match tint {
        RippleTint::ColorChange => "Color change",
        RippleTint::Current => "Current color",
        RippleTint::Custom => "Custom",
        RippleTint::Rainbow => "Rainbow",
    }
}

fn ripple_tint_tip(tint: RippleTint) -> &'static str {
    match tint {
        RippleTint::ColorChange => "The ring shifts hue as it expands. Use Color twist to control how much.",
        RippleTint::Current => "The ring only brightens the colors already on the keys from Color mode and Style.",
        RippleTint::Custom => "Paint the ring with the Ripple swatch you pick below.",
        RippleTint::Rainbow => "The ring uses its own rainbow as it travels. Color twist sets how fast the hues move.",
    }
}

fn ripple_kind_name(kind: RippleKind) -> &'static str {
    match kind {
        RippleKind::Ring => "Ring",
        RippleKind::Wave => "Wave",
        RippleKind::Pulse => "Pulse",
        RippleKind::Double => "Double",
        RippleKind::Fill => "Fill",
        RippleKind::Echo => "Echo",
    }
}

fn ripple_kind_tip(kind: RippleKind) -> &'static str {
    match kind {
        RippleKind::Ring => "A single expanding front. In 24-lamp mode it hops strip by strip.",
        RippleKind::Wave => "A thicker traveling band with a short trail behind the front.",
        RippleKind::Pulse => "A bloom that grows from the origin and fades, without a thin ring.",
        RippleKind::Double => "Two concentric rings: an outer front and a smaller inner ring.",
        RippleKind::Fill => "Fills outward from the origin, with a brighter leading edge.",
        RippleKind::Echo => "The main ring plus two delayed copies following behind.",
    }
}

fn ripple_trigger_name(trigger: RippleTrigger) -> &'static str {
    match trigger {
        RippleTrigger::All => "Any beat",
        RippleTrigger::Bass => "Bass only",
        RippleTrigger::Kick => "Kick",
        RippleTrigger::TripleKick => "Triple kick",
    }
}

fn ripple_trigger_tip(trigger: RippleTrigger) -> &'static str {
    match trigger {
        RippleTrigger::All => "Start a ring on a real onset peak (kick/snare), not every small loudness bump. Tempo cooldown stops spam.",
        RippleTrigger::Bass => "Start a ring only on a confirmed kick/bass peak. Hats and mids are ignored.",
        RippleTrigger::Kick => "Start a ring only on a kick drum (about 30–100 Hz punch). Sustained bass notes and snares are ignored.",
        RippleTrigger::TripleKick => "Wait for three confirmed kick drums in about 0.75s, then fire one ring on the third.",
    }
}

fn swipe_mode_tip(mode: SwipeMode) -> &'static str {
    match mode {
        SwipeMode::Change => "Each pass swaps in the next color.",
        SwipeMode::Fill => "Fills the keyboard, then clears it before the next pass.",
    }
}

fn stars_palette_name(p: StarsPalette) -> &'static str {
    match p {
        StarsPalette::Custom => "Custom",
        StarsPalette::White => "White",
        StarsPalette::Gold => "Gold",
        StarsPalette::Rainbow => "Rainbow",
        StarsPalette::Random => "Random",
    }
}

fn stars_palette_tip(p: StarsPalette) -> &'static str {
    match p {
        StarsPalette::Custom => "Stars take color from the four zone swatches.",
        StarsPalette::White => "Warm white stars.",
        StarsPalette::Gold => "Amber gold stars.",
        StarsPalette::Rainbow => "Each star has its own hue that slowly drifts.",
        StarsPalette::Random => "Soft mixed hues, with shooting stars picking a new color.",
    }
}

fn rain_palette_name(p: RainPalette) -> &'static str {
    match p {
        RainPalette::Ice => "Ice",
        RainPalette::Custom => "Custom",
        RainPalette::Neon => "Neon",
        RainPalette::Rainbow => "Rainbow",
    }
}

fn rain_palette_tip(p: RainPalette) -> &'static str {
    match p {
        RainPalette::Ice => "Cool blue rain.",
        RainPalette::Custom => "Drops use the four zone swatches.",
        RainPalette::Neon => "Bright green-cyan rain.",
        RainPalette::Rainbow => "Each drop keeps its own color.",
    }
}

fn aurora_palette_name(p: AuroraPalette) -> &'static str {
    match p {
        AuroraPalette::Borealis => "Borealis",
        AuroraPalette::Custom => "Custom",
        AuroraPalette::Twilight => "Twilight",
        AuroraPalette::Rainbow => "Rainbow",
    }
}

fn aurora_palette_tip(p: AuroraPalette) -> &'static str {
    match p {
        AuroraPalette::Borealis => "Classic green-to-cyan northern lights.",
        AuroraPalette::Custom => "Bands take color from the four zone swatches.",
        AuroraPalette::Twilight => "Purple and magenta dusk wash.",
        AuroraPalette::Rainbow => "Hue wanders across the full spectrum.",
    }
}

fn scanner_path_name(p: ScannerPath) -> &'static str {
    match p {
        ScannerPath::Bounce => "Bounce",
        ScannerPath::Wrap => "Wrap",
    }
}

fn scanner_path_tip(p: ScannerPath) -> &'static str {
    match p {
        ScannerPath::Bounce => "The beam turns around at each end.",
        ScannerPath::Wrap => "The beam leaves one side and comes back on the other.",
    }
}

fn scanner_palette_name(p: ScannerPalette) -> &'static str {
    match p {
        ScannerPalette::Red => "Red",
        ScannerPalette::Custom => "Custom",
        ScannerPalette::Ice => "Ice",
        ScannerPalette::Rainbow => "Rainbow",
    }
}

fn scanner_palette_tip(p: ScannerPalette) -> &'static str {
    match p {
        ScannerPalette::Red => "Classic red scanner beam.",
        ScannerPalette::Custom => "The beam uses the four zone swatches.",
        ScannerPalette::Ice => "Cyan-blue beam.",
        ScannerPalette::Rainbow => "The beam cycles hue as it travels.",
    }
}

fn battery_palette_name(p: BatteryPalette) -> &'static str {
    match p {
        BatteryPalette::Traffic => "Traffic",
        BatteryPalette::Custom => "Custom",
        BatteryPalette::Ice => "Ice",
        BatteryPalette::Heat => "Heat",
    }
}

fn battery_palette_tip(p: BatteryPalette) -> &'static str {
    match p {
        BatteryPalette::Traffic => "Green when healthy, yellow at Mid %, red at Low %.",
        BatteryPalette::Custom => "The fill samples your zone colors from left to right.",
        BatteryPalette::Ice => "Cool blue meter.",
        BatteryPalette::Heat => "Warmer color as charge rises.",
    }
}
