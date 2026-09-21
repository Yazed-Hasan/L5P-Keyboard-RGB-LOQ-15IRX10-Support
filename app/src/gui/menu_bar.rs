use crossbeam_channel::Sender;
use eframe::{
    egui::{self, Context},
    epaint::Vec2,
};
use egui_file::FileDialog;
use egui_notify::Toasts;
use std::{path::PathBuf, time::Duration};

use crate::{
    gui::modals,
    manager::{custom_effect::CustomEffect, profile::Profile},
    persist::{self, ImportedFile},
    DENY_HIDING,
};

use super::{GuiMessage, LoadedEffect};

pub struct MenuBarState {
    gui_sender: Sender<GuiMessage>,
    load_profile_dialog: FileDialog,
    load_effect_dialog: FileDialog,
    save_profile_dialog: FileDialog,
}

impl MenuBarState {
    pub(super) fn new(gui_sender: Sender<GuiMessage>) -> Self {
        let start_dir = persist::Settings::config_dir();
        let json_only: egui_file::Filter<PathBuf> = Box::new(|path: &std::path::Path| {
            path.extension()
                .and_then(|ext| ext.to_str())
                .map(|ext| ext.eq_ignore_ascii_case("json"))
                .unwrap_or(false)
        });
        Self {
            gui_sender,
            load_profile_dialog: FileDialog::open_file(Some(start_dir.clone()))
                .default_size(Vec2::splat(300.0))
                .show_files_filter(json_only),
            load_effect_dialog: FileDialog::open_file(Some(start_dir.clone())).default_size(Vec2::splat(300.0)),
            save_profile_dialog: FileDialog::save_file(Some(start_dir.join("profile.json")))
                .default_filename("profile.json")
                .default_size(Vec2::splat(300.0)),
        }
    }

    pub fn show(
        &mut self,
        ctx: &Context,
        ui: &mut egui::Ui,
        current_profile: &mut Profile,
        current_effect: &mut LoadedEffect,
        changed: &mut bool,
        toasts: &mut Toasts,
    ) -> Option<ImportedFile> {
        self.show_menu(ctx, ui, toasts);
        let imported = self.handle_load_profile(ctx, toasts);
        self.handle_save_profile(ctx, current_profile, toasts);
        self.handle_load_effect(ctx, current_effect, changed, toasts);
        imported
    }

    fn handle_load_profile(&mut self, ctx: &Context, toasts: &mut Toasts) -> Option<ImportedFile> {
        if self.load_profile_dialog.show(ctx).selected() {
            if let Some(path) = self.load_profile_dialog.path().map(|p| p.to_path_buf()) {
                self.update_paths(path.clone());
                match persist::import_from_path(&path) {
                    Ok(imported) => {
                        toasts
                            .success("Profile loaded.")
                            .duration(Some(Duration::from_millis(2500)))
                            .closable(true);
                        return Some(imported);
                    }
                    Err(err) => {
                        toasts
                            .error(err)
                            .duration(Some(Duration::from_millis(5000)))
                            .closable(true);
                    }
                }
            }
        }
        None
    }

    fn handle_save_profile(&mut self, ctx: &Context, current_profile: &mut Profile, toasts: &mut Toasts) {
        if self.save_profile_dialog.show(ctx).selected() {
            if let Some(path) = self.save_profile_dialog.path().map(|p| p.to_path_buf()) {
                match persist::export_profile(current_profile, &path) {
                    Ok(saved) => {
                        toasts
                            .success(format!("Profile saved to {}", saved.display()))
                            .duration(Some(Duration::from_millis(3500)))
                            .closable(true);
                    }
                    Err(err) => {
                        toasts.error(err).duration(Some(Duration::from_millis(5000))).closable(true);
                    }
                }
                self.update_paths(path);
            }
        }
    }

    fn handle_load_effect(&mut self, ctx: &Context, current_effect: &mut LoadedEffect, changed: &mut bool, toasts: &mut Toasts) {
        if self.load_effect_dialog.show(ctx).selected() {
            if let Some(path) = self.load_effect_dialog.path().map(|p| p.to_path_buf()) {
                match CustomEffect::from_file(&path) {
                    Ok(effect) => {
                        *current_effect = LoadedEffect::queued(effect);
                        *changed = true;
                        toasts
                            .success("Custom effect loaded.")
                            .duration(Some(Duration::from_millis(2500)))
                            .closable(true);
                    }
                    Err(_) => {
                        toasts.error("Could not load custom effect.").duration(Some(Duration::from_millis(5000))).closable(true);
                    }
                }
                self.update_paths(path);
            }
        }
    }

    fn update_paths(&mut self, path: PathBuf) {
        let mut save_paths = |path: PathBuf| {
            self.load_profile_dialog.set_path(path.clone());
            self.load_effect_dialog.set_path(path.clone());
            self.save_profile_dialog.set_path(path);
        };

        if path.exists() {
            if path.is_file() {
                if let Some(parent) = path.parent() {
                    save_paths(parent.to_path_buf())
                }
            } else {
                save_paths(path)
            }
        }
    }

    #[allow(unused_variables)]
    fn show_menu(&mut self, ctx: &Context, ui: &mut egui::Ui, toasts: &mut Toasts) {
        use egui::menu;

        menu::bar(ui, |ui| {
            ui.menu_button("Profile", |ui| {
                if ui.button("Load from file...").clicked() {
                    self.load_profile_dialog.open();
                }
                if ui.button("Save to file...").clicked() {
                    self.save_profile_dialog.open();
                }
            });

            ui.menu_button("Effect", |ui| {
                if ui.button("Open").clicked() {
                    self.load_effect_dialog.open();
                }
            });

            let about_modal = modals::about(ctx);
            if ui.button("About").clicked() {
                about_modal.open();
            }

            if ui.button("Donate").clicked() {
                open::that("https://www.buymeacoffee.com/4JXdev").unwrap();
            }

            if !*DENY_HIDING && ui.button("Exit").clicked() {
                self.gui_sender.send(GuiMessage::Quit).unwrap();
            }

            #[cfg(target_os = "windows")]
            {
                use crate::console;
                use eframe::{egui::Layout, emath::Align};
                ui.with_layout(Layout::right_to_left(Align::Center), |ui| {
                    if ui.button("📜").clicked() {
                        if !console::alloc_with_color_support() {
                            toasts.error("Could not allocate debug terminal.").duration(Some(Duration::from_millis(5000))).closable(true);
                        }
                        println!("Debug terminal enabled.");
                    }
                });
            }
        });
    }
}
