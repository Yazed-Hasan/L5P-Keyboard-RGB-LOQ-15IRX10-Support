use eframe::egui::{Color32, Context, CornerRadius, Frame, RichText, ScrollArea, Ui};
use egui_modal::Modal;

use crate::manager::{custom_effect::CustomEffect, profile::Profile};

use super::{style::SpacingStyle, LoadedEffect, State};

#[derive(Clone)]
pub struct SavedItems {
    pub custom_effects: Vec<CustomEffect>,
    pub profiles: Vec<Profile>,

    tab: Tab,
    new_item_name: String,
    request_save_as: bool,
    just_saved_profile: bool,
    just_saved_custom: bool,
}

#[derive(Copy, Clone, PartialEq, Eq)]
pub enum Tab {
    Profiles,
    CustomEffects,
}

impl SavedItems {
    pub fn new(profiles: Vec<Profile>, custom_effects: Vec<CustomEffect>) -> Self {
        Self {
            profiles,
            custom_effects,
            tab: Tab::Profiles,
            new_item_name: String::default(),
            request_save_as: false,
            just_saved_profile: false,
            just_saved_custom: false,
        }
    }

    fn setup_modal<T: Clone>(
        ctx: &Context,
        id_source: &str,
        item_name: &str,
        new_item_name: &mut String,
        items: &mut Vec<T>,
        current_item: &mut T,
        item_name_extractor: fn(&T) -> Option<String>,
        item_name_setter: fn(&mut T, String),
        saved_flag: &mut bool,
    ) -> Modal {
        let modal = Modal::new(ctx, id_source);

        modal.show(|ui| {
            modal.title(ui, item_name);
            modal.frame(ui, |ui| {
                ui.text_edit_singleline(new_item_name);
            });

            modal.buttons(ui, |ui| {
                let is_empty = new_item_name.is_empty();
                let name_not_unique = items.iter().any(|item| item_name_extractor(item) == Some(new_item_name.clone()));

                modal.button(ui, "Cancel");

                ui.add_enabled_ui(!is_empty && !name_not_unique, |ui| {
                    if modal.button(ui, "Save").clicked() {
                        item_name_setter(current_item, new_item_name.clone());
                        items.push(current_item.clone());
                        *saved_flag = true;
                    };
                });

                if is_empty {
                    ui.label("You must enter a name");
                } else if name_not_unique {
                    ui.label("Name already in use");
                }
            });
        });

        modal
    }

    pub fn setup_profile_modal(&mut self, ctx: &Context, current_profile: &mut Profile) -> Modal {
        Self::setup_modal(
            ctx,
            "profile_modal",
            "Enter the name of the profile",
            &mut self.new_item_name,
            &mut self.profiles,
            current_profile,
            |prof| prof.name.clone(),
            |prof, name| prof.name = Some(name),
            &mut self.just_saved_profile,
        )
    }

    pub fn setup_effect_modal(&mut self, ctx: &Context, loaded_effect: &mut LoadedEffect) -> Modal {
        Self::setup_modal(
            ctx,
            "effect_modal",
            "Enter the name of the custom effect",
            &mut self.new_item_name,
            &mut self.custom_effects,
            &mut loaded_effect.effect,
            |effect| effect.name.clone(),
            |effect, name| effect.name = Some(name),
            &mut self.just_saved_custom,
        )
    }

    pub fn request_save_as(&mut self) {
        self.new_item_name.clear();
        self.request_save_as = true;
    }

    pub fn upsert_named_profile(&mut self, profile: &Profile) -> bool {
        let Some(name) = profile.name.as_ref() else {
            return false;
        };
        if let Some(slot) = self.profiles.iter_mut().find(|item| item.name.as_ref() == Some(name)) {
            *slot = profile.clone();
        } else {
            self.profiles.push(profile.clone());
        }
        true
    }

    pub fn take_just_saved(&mut self) -> bool {
        let saved = self.just_saved_profile;
        self.just_saved_profile = false;
        saved
    }

    pub fn show_header(&mut self, ctx: &Context, ui: &mut Ui, current_profile: &mut Profile, loaded_effect: &mut LoadedEffect) {
        ui.selectable_value(&mut self.tab, Tab::Profiles, RichText::new("Profiles").heading());
        ui.selectable_value(&mut self.tab, Tab::CustomEffects, RichText::new("Custom Effects").heading());

        let profile_modal = self.setup_profile_modal(ctx, current_profile);
        let effect_modal = self.setup_effect_modal(ctx, loaded_effect);

        match self.tab {
            Tab::Profiles => {
                if ui.button("Save as").on_hover_text("Save the current lighting as a named profile.").clicked() {
                    self.new_item_name.clear();
                    profile_modal.open();
                }
                if self.request_save_as {
                    self.request_save_as = false;
                    self.new_item_name.clear();
                    profile_modal.open();
                }
                if ui.button("-").on_hover_text("Delete the selected named profile.").clicked() {
                    if let Some(name) = current_profile.name.clone() {
                        self.profiles.retain(|prof| prof.name.as_ref() != Some(&name));
                        current_profile.name = None;
                    }
                }
            }
            Tab::CustomEffects => {
                if loaded_effect.is_playing() && ui.button("+").clicked() {
                    self.new_item_name.clear();
                    effect_modal.open();
                }

                if ui.button("-").clicked() {
                    self.custom_effects.retain(|effect| effect != &loaded_effect.effect);
                }
            }
        }
    }

    pub fn show(&mut self, ctx: &Context, ui: &mut Ui, current_profile: &mut Profile, loaded_effect: &mut LoadedEffect, spacing: &SpacingStyle, changed: &mut bool) {
        ui.scope(|ui: &mut Ui| {
            ui.style_mut().spacing.item_spacing = spacing.default;

            ui.horizontal(|ui| {
                self.show_header(ctx, ui, current_profile, loaded_effect);
            });

            Frame {
                corner_radius: CornerRadius::same(6),
                fill: Color32::from_gray(20),
                ..Frame::default()
            }
            .show(ui, |ui| {
                ui.set_height(ui.available_height());

                ScrollArea::vertical().auto_shrink([false; 2]).show(ui, |ui| match self.tab {
                    Tab::Profiles => {
                        if self.profiles.is_empty() {
                            ui.centered_and_justified(|ui| ui.label("No profiles added"));
                        } else {
                            ui.horizontal_wrapped(|ui| {
                                let mut load_idx = None;
                                for (idx, prof) in self.profiles.iter().enumerate() {
                                    let name = prof.name.as_deref().unwrap_or("Unnamed");
                                    let selected = current_profile.name.is_some() && current_profile.name == prof.name;
                                    if ui.selectable_label(selected, name).clicked() {
                                        load_idx = Some(idx);
                                    }
                                }
                                if let Some(idx) = load_idx {
                                    *current_profile = self.profiles[idx].clone();
                                    *changed = true;
                                    loaded_effect.state = State::None;
                                }
                            });
                        }
                    }
                    Tab::CustomEffects => {
                        if self.custom_effects.is_empty() {
                            ui.centered_and_justified(|ui| ui.label("No custom effects added"));
                        } else {
                            ui.horizontal_wrapped(|ui| {
                                for effect in self.custom_effects.iter() {
                                    let name = effect.name.as_deref().unwrap_or("Unnamed");
                                    if ui.selectable_value(&mut loaded_effect.effect, effect.clone(), name).clicked() {
                                        *changed = true;
                                        loaded_effect.effect = effect.clone();
                                        loaded_effect.state = State::Queued;
                                    };
                                }
                            });
                        }
                    }
                });
            });
        });
    }
}
