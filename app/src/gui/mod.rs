use std::{process, thread, time::{Duration, Instant}};

use std::sync::{
    atomic::{AtomicBool, Ordering},
    Arc,
};

use device_query::{DeviceQuery, Keycode};
#[cfg(debug_assertions)]
use eframe::egui::style::DebugOptions;
use eframe::{
    egui::{CentralPanel, Context, CornerRadius, Frame, Layout, ScrollArea, Style, TopBottomPanel, ViewportCommand},
    emath::Align,
    epaint::{Color32, Vec2},
    CreationContext,
};

use egui_notify::Toasts;
use strum::IntoEnumIterator;
use tray_icon::menu::MenuEvent;

use crate::{
    cli::OutputType,
    enums::Effects,
    manager::{self, custom_effect::CustomEffect, profile::Profile, show_effect_ui, AudioReactParams, EffectManager, ManagerCreationError},
    persist::{ImportedFile, Settings},
    tray::{QUIT_ID, SHOW_ID},
    DENY_HIDING,
};

use self::{menu_bar::MenuBarState, saved_items::SavedItems, style::Theme};

mod menu_bar;
mod modals;
mod saved_items;
pub mod style;

pub struct App {
    instance_not_unique: bool,
    gui_tx: crossbeam_channel::Sender<GuiMessage>,
    gui_rx: crossbeam_channel::Receiver<GuiMessage>,

    has_tray: Arc<AtomicBool>,
    visible: Arc<AtomicBool>,

    manager: Option<EffectManager>,
    is_dynamic_lighting: bool,
    lamp_count: u16,
    state_changed: bool,
    loaded_effect: LoadedEffect,
    current_profile: Profile,

    menu_bar: MenuBarState,
    saved_items: SavedItems,
    mode_presets: Vec<Profile>,
    last_settings_json: String,
    pending_settings_json: String,
    last_edit: Instant,
    global_rgb: [u8; 3],
    fine_lamps: bool,
    theme: Theme,
    toasts: Toasts,
}

pub enum GuiMessage {
    CycleProfiles,
    Quit,
}

pub struct LoadedEffect {
    state: State,
    effect: CustomEffect,
}

impl LoadedEffect {
    pub fn default() -> Self {
        Self::none()
    }

    pub fn none() -> Self {
        Self {
            state: State::None,
            effect: CustomEffect::default(),
        }
    }

    pub fn queued(effect: CustomEffect) -> Self {
        Self { state: State::Queued, effect }
    }

    pub fn is_none(&self) -> bool {
        matches!(self.state, State::None)
    }

    pub fn is_queued(&self) -> bool {
        matches!(self.state, State::Queued)
    }

    pub fn is_playing(&self) -> bool {
        matches!(self.state, State::Playing)
    }
}

#[derive(Default)]
pub enum State {
    #[default]
    None,
    Queued,
    Playing,
}

impl App {
    pub fn new(output: OutputType, has_tray: Arc<AtomicBool>, visible: Arc<AtomicBool>) -> Self {
        let (gui_tx, gui_rx) = crossbeam_channel::unbounded::<GuiMessage>();

        let manager_result = EffectManager::new(manager::OperationMode::Gui);

        let instance_not_unique = if let Err(err) = &manager_result {
            &ManagerCreationError::InstanceAlreadyRunning == err.current_context()
        } else {
            false
        };

        let manager = manager_result.ok();
        let is_dynamic_lighting = manager.as_ref().map_or(false, |m| m.is_dynamic_lighting);
        let lamp_count = manager.as_ref().map_or(4, |m| m.lamp_count);

        let settings: Settings = Settings::load();
        let Settings {
            current_profile,
            profiles,
            effects,
            mode_presets,
            fine_lamps,
        } = settings;

        let gui_tx_c = gui_tx.clone();
        // Default app state
        let mut app = Self {
            instance_not_unique,
            gui_tx,
            gui_rx,

            has_tray,
            visible,

            manager,
            is_dynamic_lighting,
            lamp_count,
            // Default to true for an instant update on launch
            state_changed: true,
            loaded_effect: LoadedEffect::default(),
            current_profile,

            menu_bar: MenuBarState::new(gui_tx_c),
            saved_items: SavedItems::new(profiles, effects),
            mode_presets,
            last_settings_json: String::new(),
            pending_settings_json: String::new(),
            last_edit: Instant::now(),
            global_rgb: [0; 3],
            fine_lamps,
            theme: Theme::default(),
            toasts: Toasts::default(),
        };

        // Update the state according to the option chosen by the user
        match output {
            OutputType::Profile(profile) => app.current_profile = profile,
            OutputType::Custom(effect) => app.loaded_effect = LoadedEffect::queued(effect),
            OutputType::NoArgs => {}
            OutputType::Exit => unreachable!("Exiting the app supersedes starting the GUI"),
        }

        if let Some(manager) = &app.manager {
            manager.set_fine_lamps(app.fine_lamps);
        }

        app
    }

    pub fn init(self, cc: &CreationContext<'_>) -> Self {
        if !*DENY_HIDING {
            cc.egui_ctx.send_viewport_cmd(ViewportCommand::Visible(self.visible.load(Ordering::SeqCst)));
        }

        let egui_ctx = cc.egui_ctx.clone();
        let gui_tx = self.gui_tx.clone();
        let has_tray = self.has_tray.clone();

        std::thread::spawn(move || loop {
            if let Ok(event) = MenuEvent::receiver().recv() {
                if event.id == SHOW_ID {
                    egui_ctx.request_repaint();

                    egui_ctx.send_viewport_cmd(ViewportCommand::Visible(true));
                    egui_ctx.send_viewport_cmd(ViewportCommand::Focus);
                } else if event.id == QUIT_ID {
                    egui_ctx.request_repaint();

                    let _ = gui_tx.send(GuiMessage::Quit);
                    has_tray.store(false, Ordering::SeqCst);
                }
            }
        });

        let ctx = cc.egui_ctx.clone();
        let gui_tx_c = self.gui_tx.clone();
        if self.manager.is_some() {
            thread::spawn(move || {
                let state = device_query::DeviceState::new();
                let mut lock_switching = false;

                loop {
                    let keys = state.get_keys();

                    if keys.contains(&Keycode::LMeta) && keys.contains(&Keycode::RAlt) {
                        if !lock_switching {
                            let _ = gui_tx_c.send(GuiMessage::CycleProfiles);
                            ctx.request_repaint();
                            lock_switching = true;
                        }
                    } else {
                        lock_switching = false;
                    }

                    thread::sleep(Duration::from_millis(50));
                }
            });
        }

        self.configure_style(&cc.egui_ctx);

        self
    }
}

fn os_process_is_foreground() -> bool {
    let (our_pid, fg_pid) = foreground_pids();
    fg_pid == 0 || fg_pid == our_pid
}

fn foreground_pids() -> (u32, u32) {
    #[cfg(windows)]
    unsafe {
        use winapi::um::processthreadsapi::GetCurrentProcessId;
        use winapi::um::winuser::{GetForegroundWindow, GetWindowThreadProcessId};
        let our_pid = GetCurrentProcessId();
        let hwnd = GetForegroundWindow();
        if hwnd.is_null() {
            return (our_pid, 0);
        }
        let mut fg_pid = 0u32;
        GetWindowThreadProcessId(hwnd, &mut fg_pid);
        (our_pid, fg_pid)
    }
    #[cfg(not(windows))]
    {
        (0, 0)
    }
}

impl eframe::App for App {
    fn update(&mut self, ctx: &eframe::egui::Context, _frame: &mut eframe::Frame) {
        if let Some(manager) = &self.manager {
            // Only the OS foreground PID. Viewport/hover flaps were switching
            // WDL and HID against each other when clicking away.
            let window_active = os_process_is_foreground();
            static LAST_ACTIVE: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(true);
            let prev = LAST_ACTIVE.swap(window_active, Ordering::Relaxed);
            if prev != window_active {
                let (our_pid, fg_pid) = foreground_pids();
                let viewport = ctx.input(|i| i.viewport().focused);
                legion_rgb_driver::debug_log(&format!(
                    "GUI: window_active {} -> {} our_pid={} fg_pid={} viewport={:?}",
                    prev, window_active, our_pid, fg_pid, viewport
                ));
            }
            manager.window_active.store(window_active, Ordering::Relaxed);
        }

        if let Ok(message) = self.gui_rx.try_recv() {
            match message {
                GuiMessage::CycleProfiles => self.cycle_profiles(),
                GuiMessage::Quit => self.exit_app(),
            }
        }

        // Show active toast messages
        self.toasts.show(ctx);

        if *DENY_HIDING && !self.visible.load(Ordering::SeqCst) {
            self.visible.store(true, Ordering::SeqCst);
            self.toasts
                .warning("Window hiding is currently not supported.\nSee https://github.com/4JX/L5P-Keyboard-RGB/issues/181")
                .duration(None);
        }

        if self.instance_not_unique && modals::unique_instance(ctx) {
            self.exit_app();
        }

        if !self.instance_not_unique && self.manager.is_none() && modals::manager_error(ctx) {
            self.exit_app();
        }

        TopBottomPanel::top("top-panel").show(ctx, |ui| {
            if let Some(imported) = self.menu_bar.show(
                ctx,
                ui,
                &mut self.current_profile,
                &mut self.loaded_effect,
                &mut self.state_changed,
                &mut self.toasts,
            ) {
                self.import_profiles(imported);
            }
        });

        CentralPanel::default()
            .frame(Frame::new().inner_margin(self.theme.spacing.large).fill(Color32::from_gray(26)))
            .show(ctx, |ui| {
                ui.style_mut().spacing.item_spacing = Vec2::splat(self.theme.spacing.large);
                self.show_ui_elements(ctx, ui);
            });

        if self.state_changed {
            self.update_state();
        }

        self.maybe_autosave();

        self.handle_close_request(ctx);
    }

    fn on_exit(&mut self, _gl: Option<&eframe::glow::Context>) {
        self.persist_now(false);

        self.visible.store(false, Ordering::SeqCst);

        if let Some(manager) = self.manager.take() {
            manager.shutdown();
        }
    }
}

impl App {
    fn configure_style(&self, ctx: &Context) {
        let mut style = Style {
            visuals: self.theme.visuals.clone(),
            #[cfg(debug_assertions)]
            debug: DebugOptions {
                debug_on_hover: false,
                debug_on_hover_with_all_modifiers: false,
                hover_shows_next: false,
                show_expand_width: false,
                show_expand_height: false,
                show_resize: false,
                show_interactive_widgets: false,
                show_widget_hits: false,
                show_unaligned: false,
            },
            ..Style::default()
        };
        style.interaction.tooltip_delay = 0.0;
        style.interaction.show_tooltips_only_when_still = false;
        style.spacing.tooltip_width = 280.0;

        ctx.set_style(style);
    }

    fn exit_app(&mut self) {
        use eframe::App;

        self.on_exit(None);

        process::exit(0);
    }

    fn cycle_profiles(&mut self) {
        let len = self.saved_items.profiles.len();

        let current_profile_name = &self.current_profile.name;

        if let Some((i, _)) = self.saved_items.profiles.iter().enumerate().find(|(_, profile)| &profile.name == current_profile_name) {
            if i == len - 1 && len > 0 {
                self.current_profile = self.saved_items.profiles[0].clone();
            } else {
                self.current_profile = self.saved_items.profiles[i + 1].clone();
            }

            self.state_changed = true;
        }
    }

    fn collect_settings(&self) -> Settings {
        let mut mode_presets = self.mode_presets.clone();
        let mut snap = self.current_profile.clone();
        snap.name = None;
        if let Some(slot) = mode_presets.iter_mut().find(|preset| preset.effect == snap.effect) {
            *slot = snap;
        } else {
            mode_presets.push(snap);
        }
        Settings::new(
            self.saved_items.profiles.clone(),
            self.saved_items.custom_effects.clone(),
            self.current_profile.clone(),
            mode_presets,
            self.fine_lamps,
        )
    }

    fn persist_now(&mut self, toast: bool) {
        let settings = self.collect_settings();
        self.mode_presets = settings.mode_presets.clone();
        settings.save();
        let json = serde_json::to_string(&settings).unwrap_or_default();
        self.last_settings_json = json.clone();
        self.pending_settings_json = json;
        if toast {
            self.toasts
                .success("Settings saved.")
                .duration(Some(Duration::from_millis(2500)))
                .closable(true);
        }
    }

    fn maybe_autosave(&mut self) {
        let json = serde_json::to_string(&self.collect_settings()).unwrap_or_default();
        if json == self.last_settings_json {
            return;
        }
        if json != self.pending_settings_json {
            self.pending_settings_json = json;
            self.last_edit = Instant::now();
            return;
        }
        if self.last_edit.elapsed() >= Duration::from_millis(700) {
            self.persist_now(false);
        }
    }

    fn save_everything(&mut self) {
        if self.current_profile.name.is_some() {
            self.saved_items.upsert_named_profile(&self.current_profile);
            let name = self.current_profile.name.clone().unwrap_or_default();
            self.persist_now(false);
            self.toasts
                .success(format!("Profile \"{name}\" saved."))
                .duration(Some(Duration::from_millis(2500)))
                .closable(true);
        } else {
            self.saved_items.request_save_as();
        }
    }

    fn apply_profile(&mut self, profile: Profile) {
        self.current_profile = profile;
        self.store_mode_preset();
        self.loaded_effect.state = State::None;
        self.state_changed = true;
    }

    fn import_profiles(&mut self, imported: ImportedFile) {
        match imported {
            ImportedFile::Profile(profile) => {
                if profile.name.is_some() {
                    self.saved_items.upsert_named_profile(&profile);
                }
                self.apply_profile(profile);
            }
            ImportedFile::Bundle(settings) => {
                for profile in settings.profiles {
                    self.saved_items.upsert_named_profile(&profile);
                }
                if settings.current_profile.name.is_some() {
                    self.saved_items.upsert_named_profile(&settings.current_profile);
                }
                if !settings.mode_presets.is_empty() {
                    self.mode_presets = settings.mode_presets;
                }
                self.apply_profile(settings.current_profile);
            }
        }
        self.persist_now(false);
    }

    fn store_mode_preset(&mut self) {
        let mut snap = self.current_profile.clone();
        snap.name = None;
        if let Some(slot) = self.mode_presets.iter_mut().find(|preset| preset.effect == snap.effect) {
            *slot = snap;
        } else {
            self.mode_presets.push(snap);
        }
    }

    fn switch_to_effect(&mut self, factory: Effects) {
        if self.current_profile.effect == factory {
            return;
        }
        self.store_mode_preset();
        let name = self.current_profile.name.clone();
        let brightness = self.current_profile.brightness;
        let brightness_level = self.current_profile.brightness_level;
        if let Some(preset) = self.mode_presets.iter().find(|preset| preset.effect == factory).cloned() {
            self.current_profile = preset;
        } else {
            self.current_profile.effect = factory;
            if matches!(factory, Effects::AudioReact { .. }) && self.current_profile.rgb_zones.iter().all(|zone| zone.rgb == [0, 0, 0])
            {
                self.current_profile.rgb_zones = crate::manager::profile::arr_to_zones(crate::manager::profile::DEFAULT_AUDIO_ZONE_RGB);
            }
        }
        self.current_profile.name = name;
        self.current_profile.brightness = brightness;
        self.current_profile.brightness_level = brightness_level;
    }

    fn show_ui_elements(&mut self, ctx: &Context, ui: &mut eframe::egui::Ui) {
        ui.with_layout(Layout::left_to_right(Align::Center).with_cross_justify(true), |ui| {
            ui.vertical(|ui| {
                if self.lamp_count >= 8 {
                    let text = "Off uses 4 zones. On paints all 24 lamp columns left to right (this keyboard has no independent rows).";
                    ui.horizontal(|ui| {
                        if ui.checkbox(&mut self.fine_lamps, "24-lamp mode").on_hover_text(text).changed() {
                            if let Some(manager) = &self.manager {
                                manager.set_fine_lamps(self.fine_lamps);
                            }
                            self.state_changed = true;
                        }
                        if ui.small_button("?").on_hover_text(text).clicked() {}
                    });
                }

                let is_audio = matches!(self.current_profile.effect, crate::enums::Effects::AudioReact { .. });
                let is_scene = self.current_profile.effect.is_scene();
                let live_colors = is_audio || is_scene;
                let can_tweak_colors = self.current_profile.effect.takes_color_array() && self.loaded_effect.is_none();
                let lamp_count = self.lamp_count.max(1) as usize;
                let show_columns = self.fine_lamps
                    && can_tweak_colors
                    && !is_audio
                    && lamp_count >= 8
                    && matches!(self.current_profile.effect, Effects::Static | Effects::Breath);

                let res = ui.add_enabled_ui(can_tweak_colors, |ui| {
                    ui.style_mut().spacing.item_spacing = Vec2::splat(self.theme.spacing.medium);
                    let response = ui.horizontal(|ui| {
                        ui.style_mut().spacing.interact_size = Vec2::new(70.0, 50.0);

                        for i in 0..4 {
                            let changed = ui
                                .color_edit_button_srgb(&mut self.current_profile.rgb_zones[i].rgb)
                                .on_hover_text(match i {
                                    0 => "Zone 1 (left). Fills the left group of lamp columns. In Audio React this is usually bass.",
                                    1 => "Zone 2. Fills the next group of lamp columns. In Audio React this is usually low-mids.",
                                    2 => "Zone 3. Fills the next group of lamp columns. In Audio React this is usually treble.",
                                    _ => "Zone 4 (right). Fills the right group of lamp columns. In Audio React this is usually presence/air.",
                                })
                                .on_disabled_hover_text("These colors apply when the effect uses custom zone colors.")
                                .changed();
                            if changed {
                                if show_columns {
                                    self.current_profile.fill_zone_lamps(i, lamp_count);
                                }
                                if !live_colors {
                                    self.state_changed = true;
                                }
                            }
                        }
                    });

                    ui.style_mut().spacing.item_spacing = Vec2::splat(4.0);
                    ui.style_mut().spacing.interact_size = Vec2::new(response.response.rect.width(), 30.0);
                    if ui
                        .color_edit_button_srgb(&mut self.global_rgb)
                        .on_hover_text("Set every zone and lamp column to the same color.")
                        .on_disabled_hover_text("These colors apply when the effect uses custom zone colors.")
                        .changed()
                    {
                        for i in 0..4 {
                            self.current_profile.rgb_zones[i].rgb = self.global_rgb;
                        }
                        if show_columns {
                            self.current_profile.fill_all_lamps(self.global_rgb, lamp_count);
                        }
                        if !live_colors {
                            self.state_changed = true;
                        }
                    }

                    if show_columns {
                        ui.add_space(6.0);
                        ui.label("Lamp columns (same as Legion Space custom theme)");
                        self.current_profile.ensure_lamps(lamp_count);
                        ui.style_mut().spacing.interact_size = Vec2::new(18.0, 28.0);
                        ui.style_mut().spacing.item_spacing = Vec2::new(2.0, 2.0);
                        ui.horizontal_wrapped(|ui| {
                            for i in 0..lamp_count {
                                let o = i * 3;
                                if o + 2 >= self.current_profile.lamp_rgb.len() {
                                    break;
                                }
                                let mut color = [
                                    self.current_profile.lamp_rgb[o],
                                    self.current_profile.lamp_rgb[o + 1],
                                    self.current_profile.lamp_rgb[o + 2],
                                ];
                                let response = ui.push_id(i, |ui| {
                                    ui.color_edit_button_srgb(&mut color)
                                        .on_hover_text(format!("Column {} of {}", i + 1, lamp_count))
                                });
                                if response.inner.changed() {
                                    self.current_profile.set_lamp(i, color, lamp_count);
                                    self.state_changed = true;
                                }
                            }
                        });
                    }

                    response.response
                });

                ui.set_width(res.inner.rect.width());

                ui.horizontal(|ui| {
                    if ui
                        .button("Save")
                        .on_hover_text("Save current lighting. If a named profile is selected, that profile is updated too.")
                        .clicked()
                    {
                        self.save_everything();
                    }
                    if ui
                        .button("Save as")
                        .on_hover_text("Save the current lighting as a new named profile.")
                        .clicked()
                    {
                        self.saved_items.request_save_as();
                    }
                    if ui
                        .button("Reset colors")
                        .on_hover_text("Restore this mode's default zone and column colors.")
                        .clicked()
                    {
                        if matches!(self.current_profile.effect, Effects::AudioReact { .. }) {
                            self.current_profile.rgb_zones =
                                crate::manager::profile::arr_to_zones(crate::manager::profile::DEFAULT_AUDIO_ZONE_RGB);
                        } else {
                            self.current_profile.rgb_zones = crate::manager::profile::arr_to_zones([255; 12]);
                        }
                        self.current_profile.clear_lamp_colors();
                        if self.lamp_count >= 8 {
                            self.current_profile.ensure_lamps(self.lamp_count as usize);
                        }
                        self.state_changed = true;
                    }
                });

                self.show_effect_ui(ui);

                self.saved_items
                    .show(ctx, ui, &mut self.current_profile, &mut self.loaded_effect, &self.theme.spacing, &mut self.state_changed);
                if self.saved_items.take_just_saved() {
                    self.store_mode_preset();
                    let name = self.current_profile.name.clone().unwrap_or_else(|| "profile".to_string());
                    self.persist_now(false);
                    self.toasts
                        .success(format!("Profile \"{name}\" saved."))
                        .duration(Some(Duration::from_millis(2500)))
                        .closable(true);
                }
            });

            ui.vertical_centered_justified(|ui| {
                if self.loaded_effect.is_playing() && ui.button("Stop custom effect").clicked() {
                    self.loaded_effect.state = State::None;
                    self.state_changed = true;
                }

                Frame {
                    corner_radius: CornerRadius::same(6),
                    fill: Color32::from_gray(20),
                    ..Frame::default()
                }
                .show(ui, |ui| {
                    ui.style_mut().spacing.item_spacing = self.theme.spacing.default;
                    ScrollArea::vertical().show(ui, |ui| {
                        ui.with_layout(Layout::top_down_justified(Align::Min), |ui| {
                            for val in Effects::iter() {
                                let factory = val.factory_default();
                                let text: &'static str = factory.into();
                                let selected = self.current_profile.effect == factory;
                                if ui
                                    .selectable_label(selected, text)
                                    .on_hover_text(effect_hover_tip(&factory))
                                    .on_disabled_hover_text(effect_hover_tip(&factory))
                                    .clicked()
                                {
                                    self.switch_to_effect(factory);
                                    self.state_changed = true;
                                    self.loaded_effect.state = State::None;
                                }
                            }
                        });
                    });
                });
            });
        });
    }

    fn show_effect_ui(&mut self, ui: &mut eframe::egui::Ui) {
        ui.add_enabled_ui(self.loaded_effect.is_none(), |ui| {
            let mut live_speed = None;
            let hud = self.manager.as_ref().map(|m| m.audio_hud()).unwrap_or_default();
            show_effect_ui(
                ui,
                &mut self.current_profile,
                &mut self.state_changed,
                &self.theme,
                self.is_dynamic_lighting,
                &mut live_speed,
                hud,
            );
            if let Some(speed) = live_speed {
                if let Some(manager) = &self.manager {
                    manager.set_live_speed(speed);
                }
            }
            if matches!(self.current_profile.effect, crate::enums::Effects::AudioReact { .. }) {
                if let Some(manager) = &self.manager {
                    manager.set_live_audio(
                        AudioReactParams::from_effect(self.current_profile.effect).with_rgb(self.current_profile.rgb_array()),
                    );
                }
            }
            if self.current_profile.effect.is_scene() {
                if let Some(manager) = &self.manager {
                    manager.set_live_scene(self.current_profile.effect, self.current_profile.rgb_array());
                }
            }
        });
    }

    fn update_state(&mut self) {
        if let Some(manager) = self.manager.as_mut() {
            if self.loaded_effect.is_none() {
                manager.set_profile(self.current_profile.clone());
            } else if self.loaded_effect.is_queued() {
                self.loaded_effect.state = State::Playing;

                let effect = self.loaded_effect.effect.clone();
                manager.custom_effect(effect);
            }
        }

        self.state_changed = false;
    }

    fn handle_close_request(&mut self, ctx: &Context) {
        if ctx.input(|i| i.viewport().close_requested()) && !*DENY_HIDING {
            if self.has_tray.load(Ordering::Relaxed) {
                ctx.send_viewport_cmd(ViewportCommand::CancelClose);
                ctx.send_viewport_cmd(ViewportCommand::Visible(false));
            } else {
                // Close normally
            }
        }
    }
}

fn effect_hover_tip(effect: &Effects) -> &'static str {
    match effect {
        Effects::Static => "Solid colors. No animation.",
        Effects::Breath => "Fades the colors in and out.",
        Effects::Smooth => "Soft color cycling across the zones.",
        Effects::Wave => "A wave of color that travels left or right.",
        Effects::Lightning => "Random lightning-style flashes.",
        Effects::AmbientLight { .. } => "Copies colors from your screen onto the keyboard.",
        Effects::SmoothWave { .. } => "A smoother traveling wave. Swipe mode changes how it fills.",
        Effects::Swipe { .. } => "Colors sweep across the zones.",
        Effects::Disco => "Quick random zone flashes.",
        Effects::Christmas => "Red and green holiday pattern.",
        Effects::Fade => "Cross-fades between the zone colors.",
        Effects::Temperature => "Color follows CPU temperature.",
        Effects::Ripple => "Waves spread from the keys you press.",
        Effects::AudioReact { .. } => {
            "Lights follow the sound this PC is playing. Hover each Audio React slider for what it does."
        }
        Effects::Stars { .. } => "Slow twinkles on a dark field. Not a party strobe like Disco.",
        Effects::Rain { .. } => "Drops travel across the keys with short trails and a splash at the end.",
        Effects::Aurora { .. } => "Overlapping northern-lights bands that drift slowly.",
        Effects::Scanner { .. } => "A bouncing hotspot with a trail, like a scanner bar.",
        Effects::Battery { .. } => "A left-to-right charge meter. Pulses while plugged in.",
        Effects::TypeHeat { .. } => "Zones heat up as you type and cool when idle. Not a spreading wave like Ripple.",
        Effects::Pacifica { .. } => "Layered blue-green ocean sines. Not raindrops and not aurora bands.",
        Effects::DigitalRain { .. } => "Heads travel across the keyboard with a fading trail. In 24-lamp mode they use all 24 columns.",
        Effects::Fireworks { .. } => "Random bursts that pop and fade. No audio needed.",
        Effects::Nexus { .. } => "A pulse on the keys you press: that quarter of columns lights, with a little bleed into the neighbors. Not a spreading ring like Ripple.",
        Effects::Comet { .. } => "A meteor with a fat head and a long tail. Head, glow, fade, sparkle, and extra comets are on the sliders. Wrap flies off the edge; Bounce turns around. Not the short Scanner beam.",
        Effects::Juggle { .. } => "Several colored dots weave back and forth with trails, like WLED Juggle. Not a single Scanner beam and not a Comet meteor.",
        Effects::Bounce { .. } => "Balls fall with gravity and bounce at the ends. Physics bounce, not Comet's smooth meteor turn.",
        Effects::Dissolve { .. } => "Keys fill in a random order, pause, then melt away. Best in 24-lamp mode.",
    }
}
