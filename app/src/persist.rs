use std::{
    env,
    fs::{self, File},
    io::Write,
    path::PathBuf,
};

use crate::manager::{custom_effect::CustomEffect, profile::Profile};
use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize, Default, Clone)]
pub struct Settings {
    pub profiles: Vec<Profile>,
    pub effects: Vec<CustomEffect>,
    // Up to 0.19.5
    #[serde(alias = "ui_state")]
    pub current_profile: Profile,
    /// Last-used colors and params for each lighting mode, so switching modes does not wipe them.
    #[serde(default)]
    pub mode_presets: Vec<Profile>,
    /// Paint every lamp strip instead of collapsing to 4 zones. Off is lighter.
    #[serde(default)]
    pub fine_lamps: bool,
}

impl Settings {
    pub fn new(profiles: Vec<Profile>, effects: Vec<CustomEffect>, current_profile: Profile, mode_presets: Vec<Profile>, fine_lamps: bool) -> Self {
        Self {
            profiles,
            effects,
            current_profile,
            mode_presets,
            fine_lamps,
        }
    }

    /// Load the settings from the configured path or generate default ones if an error occurs
    pub fn load() -> Self {
        for path in Self::candidate_paths() {
            if let Ok(string) = fs::read_to_string(&path) {
                match serde_json::from_str::<Self>(&string) {
                    Ok(persist) => {
                        legion_rgb_driver::debug_log(&format!("SETTINGS: loaded {}", path.display()));
                        return persist;
                    }
                    Err(err) => {
                        legion_rgb_driver::debug_log(&format!("SETTINGS: could not parse {}: {err}", path.display()));
                    }
                }
            }
        }

        legion_rgb_driver::debug_log("SETTINGS: using defaults");
        Self::default()
    }

    /// Save the settings to the configured path
    pub fn save(&self) {
        let path = Self::get_location();
        if let Some(parent) = path.parent() {
            let _ = fs::create_dir_all(parent);
        }
        match serde_json::to_string_pretty(self) {
            Ok(stringified_json) => {
                let tmp = path.with_extension("json.tmp");
                match File::create(&tmp).and_then(|mut file| file.write_all(stringified_json.as_bytes())) {
                    Ok(()) => {
                        let _ = fs::remove_file(&path);
                        if fs::rename(&tmp, &path).is_err() {
                            let _ = fs::copy(&tmp, &path);
                            let _ = fs::remove_file(&tmp);
                        }
                        legion_rgb_driver::debug_log(&format!("SETTINGS: saved {}", path.display()));
                    }
                    Err(err) => {
                        legion_rgb_driver::debug_log(&format!("SETTINGS: write failed {}: {err}", path.display()));
                    }
                }
            }
            Err(err) => {
                legion_rgb_driver::debug_log(&format!("SETTINGS: serialize failed: {err}"));
            }
        }
    }

    fn candidate_paths() -> Vec<PathBuf> {
        let mut paths = Vec::new();
        let primary = Self::get_location();
        paths.push(primary.clone());
        let cwd = PathBuf::from("./settings.json");
        if cwd.canonicalize().ok() != primary.canonicalize().ok() {
            paths.push(cwd);
        }
        paths
    }

    fn get_location() -> PathBuf {
        if let Ok(maybe_path) = env::var("LEGION_KEYBOARD_CONFIG") {
            let path = PathBuf::from(maybe_path);
            if !path.as_os_str().is_empty() {
                return path;
            }
        }

        if let Ok(exe) = env::current_exe() {
            if let Some(dir) = exe.parent() {
                return dir.join("settings.json");
            }
        }

        PathBuf::from("./settings.json")
    }

    pub fn config_dir() -> PathBuf {
        Self::get_location()
            .parent()
            .map(|path| path.to_path_buf())
            .unwrap_or_else(|| PathBuf::from("."))
    }
}

pub enum ImportedFile {
    Profile(Profile),
    Bundle(Settings),
}

pub fn with_json_ext(path: &std::path::Path) -> PathBuf {
    match path.extension().and_then(|ext| ext.to_str()) {
        Some(ext) if ext.eq_ignore_ascii_case("json") => path.to_path_buf(),
        _ => path.with_extension("json"),
    }
}

pub fn import_from_path(path: &std::path::Path) -> Result<ImportedFile, String> {
    let text = fs::read_to_string(path).map_err(|err| format!("Could not read file: {err}"))?;
    if let Ok(profile) = serde_json::from_str::<Profile>(&text) {
        return Ok(ImportedFile::Profile(profile));
    }
    if let Ok(settings) = serde_json::from_str::<Settings>(&text) {
        return Ok(ImportedFile::Bundle(settings));
    }
    Err("This file is not a keyboard profile or settings export.".to_string())
}

pub fn export_profile(profile: &Profile, path: &std::path::Path) -> Result<PathBuf, String> {
    let path = with_json_ext(path);
    let json = serde_json::to_string_pretty(profile).map_err(|err| format!("Could not write profile: {err}"))?;
    fs::write(&path, json).map_err(|err| format!("Could not write profile: {err}"))?;
    Ok(path)
}
