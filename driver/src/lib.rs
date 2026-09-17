use error::{RangeError, RangeErrorKind, Result};
use hidapi::{HidApi, HidDevice};
use std::{
    collections::BTreeMap,
    collections::BTreeSet,
    io::Write,
    sync::{
        atomic::{AtomicBool, AtomicU64, Ordering},
        Arc,
    },
    thread,
    time::Duration,
};

pub mod error;

/// Append a timestamped line to "legion_rgb_debug.log" next to the executable.
pub fn debug_log(msg: &str) {
    log_to_file(msg);
}

fn log_to_file(msg: &str) {
    if reverse_engineering_mode_enabled() {
        eprintln!("[legion-rgb] {}", msg);
    }

    let path = std::env::current_exe()
        .unwrap_or_default()
        .parent()
        .unwrap_or(std::path::Path::new("."))
        .join("legion_rgb_debug.log");
    if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true).open(&path) {
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default();
        let _ = writeln!(f, "[{:.3}] {}", now.as_secs_f64(), msg);
    }
}

fn reverse_engineering_mode_enabled() -> bool {
    std::env::var("LEGION_RGB_REVERSE_MODE")
        .map(|value| {
            let value = value.trim().to_ascii_lowercase();
            value == "1" || value == "true" || value == "yes" || value == "on"
        })
        .unwrap_or(false)
}

fn force_vendor_only_enabled() -> bool {
    std::env::var("LEGION_RGB_FORCE_VENDOR_ONLY")
        .map(|value| {
            let value = value.trim().to_ascii_lowercase();
            value == "1" || value == "true" || value == "yes" || value == "on"
        })
        .unwrap_or(false)
}

fn loq_vendor_primary_enabled() -> bool {
    // For LOQ 15IRX10, prefer the MI_00 vendor endpoint by default.
    // Set LEGION_RGB_LOQ_USE_LAMPARRAY=1 to re-enable mixed LampArray writes.
    !std::env::var("LEGION_RGB_LOQ_USE_LAMPARRAY")
        .map(|value| {
            let value = value.trim().to_ascii_lowercase();
            value == "1" || value == "true" || value == "yes" || value == "on"
        })
        .unwrap_or(false)
}

#[cfg(target_os = "windows")]
fn loq_hid_first_enabled() -> bool {
    // Off by default. On LOQ 15IRX10 only WinRT actually updates the keys.
    // HID-first opens successfully but leaves the keyboard dark.
    std::env::var("LEGION_RGB_LOQ_HID_FIRST")
        .map(|value| {
            let value = value.trim().to_ascii_lowercase();
            value == "1" || value == "true" || value == "yes" || value == "on"
        })
        .unwrap_or(false)
}

#[cfg(target_os = "windows")]
fn wdl_hid_fallback_enabled(_is_loq_15irx10: bool) -> bool {
    // Opening the LampArray HID while WinRT owns it can stop focused writes.
    std::env::var("LEGION_RGB_WDL_USE_HID_FALLBACK")
        .map(|value| {
            let value = value.trim().to_ascii_lowercase();
            value == "1" || value == "true" || value == "yes" || value == "on"
        })
        .unwrap_or(false)
}

#[cfg(target_os = "windows")]
fn wdl_vendor_fallback_enabled(_is_loq_15irx10: bool) -> bool {
    // Vendor SAVE_PROFILE fights WinRT LampArray and flickers the keys.
    // Keep it opt-in; Windows hold lighting is the unfocused path.
    std::env::var("LEGION_RGB_WDL_USE_VENDOR_FALLBACK")
        .map(|value| {
            let value = value.trim().to_ascii_lowercase();
            value == "1" || value == "true" || value == "yes" || value == "on"
        })
        .unwrap_or(false)
}

fn log_hid_inventory(api: &HidApi, context: &str) {
    if !reverse_engineering_mode_enabled() {
        return;
    }

    log_to_file(&format!("RE-MODE [{}]: HID inventory begin", context));
    for d in api.device_list() {
        if d.vendor_id() != 0x048d {
            continue;
        }

        log_to_file(&format!(
            "RE-HID [{}]: vid={:#06x} pid={:#06x} usage_page={:#06x} usage={:#06x} iface={} path={}",
            context,
            d.vendor_id(),
            d.product_id(),
            d.usage_page(),
            d.usage(),
            d.interface_number(),
            d.path().to_string_lossy()
        ));
    }
    log_to_file(&format!("RE-MODE [{}]: HID inventory end", context));
}

fn log_descriptor_preview(context: &str, descriptor: &[u8]) {
    if !reverse_engineering_mode_enabled() {
        return;
    }

    let preview_len = std::cmp::min(96, descriptor.len());
    let preview = descriptor[..preview_len]
        .iter()
        .map(|b| format!("{:02x}", b))
        .collect::<Vec<_>>()
        .join(" ");

    log_to_file(&format!(
        "RE-DESC [{}]: len={} preview_len={} bytes={}{}",
        context,
        descriptor.len(),
        preview_len,
        preview,
        if descriptor.len() > preview_len { " ..." } else { "" }
    ));
}

#[derive(Clone, Debug, Default)]
struct HidReportIdSets {
    output: Vec<u8>,
    feature: Vec<u8>,
    output_lengths: BTreeMap<u8, usize>,
    feature_lengths: BTreeMap<u8, usize>,
}

fn parse_report_id_sets(descriptor: &[u8]) -> HidReportIdSets {
    let mut output = BTreeSet::new();
    let mut feature = BTreeSet::new();
    let mut output_lengths = BTreeMap::new();
    let mut feature_lengths = BTreeMap::new();
    let mut report_id: u8 = 0;
    let mut report_size_bits: usize = 0;
    let mut report_count: usize = 0;

    let mut i = 0;
    while i < descriptor.len() {
        let prefix = descriptor[i];

        if prefix == 0xFE {
            if i + 2 >= descriptor.len() {
                break;
            }
            let data_size = descriptor[i + 1] as usize;
            i += 3 + data_size;
            continue;
        }

        let size = match prefix & 0x03 {
            0 => 0usize,
            1 => 1,
            2 => 2,
            3 => 4,
            _ => unreachable!(),
        };
        let item_type = (prefix >> 2) & 0x03;
        let tag = (prefix >> 4) & 0x0F;

        if i + 1 + size > descriptor.len() {
            break;
        }

        let value = match size {
            1 => descriptor[i + 1] as u32,
            2 => u16::from_le_bytes([descriptor[i + 1], descriptor[i + 2]]) as u32,
            4 => u32::from_le_bytes([
                descriptor[i + 1],
                descriptor[i + 2],
                descriptor[i + 3],
                descriptor[i + 4],
            ]),
            _ => 0,
        };

        match item_type {
            0 => match tag {
                9 => {
                    output.insert(report_id);
                    let byte_len = (report_size_bits * report_count).div_ceil(8) + 1;
                    output_lengths
                        .entry(report_id)
                        .and_modify(|len: &mut usize| *len = (*len).max(byte_len))
                        .or_insert(byte_len);
                }
                11 => {
                    feature.insert(report_id);
                    let byte_len = (report_size_bits * report_count).div_ceil(8) + 1;
                    feature_lengths
                        .entry(report_id)
                        .and_modify(|len: &mut usize| *len = (*len).max(byte_len))
                        .or_insert(byte_len);
                }
                _ => {}
            },
            1 => {
                match tag {
                    7 => report_size_bits = value as usize,
                    8 => report_id = value as u8,
                    9 => report_count = value as usize,
                    _ => {}
                }
            }
            _ => {}
        }

        i += 1 + size;
    }

    HidReportIdSets {
        output: output.into_iter().collect(),
        feature: feature.into_iter().collect(),
        output_lengths,
        feature_lengths,
    }
}

fn log_report_id_sets(context: &str, descriptor: &[u8]) -> HidReportIdSets {
    let ids = parse_report_id_sets(descriptor);
    if reverse_engineering_mode_enabled() {
        log_to_file(&format!(
            "RE-DESC [{}]: output_report_ids={:?} feature_report_ids={:?} output_lengths={:?} feature_lengths={:?}",
            context, ids.output, ids.feature, ids.output_lengths, ids.feature_lengths
        ));
    }
    ids
}

fn log_feature_report_snapshot(device: &HidDevice, context: &str, report_id: u8, report_len: usize) {
    if !reverse_engineering_mode_enabled() {
        return;
    }

    let safe_len = report_len.clamp(2, 960);
    let mut buf = vec![0u8; safe_len];
    buf[0] = report_id;
    match device.get_feature_report(&mut buf) {
        Ok(n) => {
            let preview = buf[..n]
                .iter()
                .map(|b| format!("{:02x}", b))
                .collect::<Vec<_>>()
                .join(" ");
            log_to_file(&format!(
                "RE-FEATURE [{}]: report_id={:#04x} requested_len={} actual_len={} bytes={}",
                context, report_id, safe_len, n, preview
            ));
        }
        Err(e) => {
            log_to_file(&format!(
                "RE-FEATURE [{}]: report_id={:#04x} requested_len={} failed: {}",
                context, report_id, safe_len, e
            ));
        }
    }
}

fn is_vivid_rgb_frame(rgb: &[u8; 12]) -> bool {
    rgb.iter().any(|v| *v >= 24)
}

#[cfg(target_os = "windows")]
fn should_disable_windows_ambient_lighting_for_loq() -> bool {
    std::env::var("LEGION_RGB_DISABLE_WDL_AMBIENT")
        .ok()
        .map(|v| v == "1")
        .unwrap_or(true)
}

/// Keep Windows lighting available as the unfocused painter, and stop Legion
/// from sitting above the keyboard. WinRT is focus-gated; firmware save is not
/// enough on this LOQ, so Windows' own solid/gradient effect has to take over
/// the moment this window is not foreground.
#[cfg(target_os = "windows")]
fn disable_windows_dynamic_lighting_ambient() {
    prepare_unfocused_lighting(true);
}

#[cfg(target_os = "windows")]
fn prepare_unfocused_lighting(verbose: bool) {
    disable_legion_lighting_background_apps(verbose);
    stop_legion_lighting_helpers();
    restore_windows_lighting_provider(r"HKCU\Software\Microsoft\Lighting\Providers", verbose);
    remove_legion_lighting_providers(r"HKCU\Software\Microsoft\Lighting\Providers");
    // Keep Windows hold lighting OFF while we own the LampArray so the two
    // writers cannot flicker against each other.
    if reg_add_dword(r"HKCU\Software\Microsoft\Lighting", "AmbientLightingEnabled", 0) && verbose {
        log_to_file("WDL: Windows hold lighting off while this app is in control");
    }
    let _ = reg_add_dword(r"HKCU\Software\Microsoft\Lighting", "UseSystemAccentColor", 0);
    set_controlled_by_foreground_app(true);
    for device_key in reg_query_subkeys(r"HKCU\Software\Microsoft\Lighting\Devices") {
        let _ = reg_add_dword(&device_key, "AmbientLightingEnabled", 0);
        let _ = reg_add_dword(&device_key, "UseSystemAccentColor", 0);
        restore_windows_lighting_provider(&format!("{}\\Providers", device_key), verbose);
        remove_legion_lighting_providers(&format!("{}\\Providers", device_key));
    }
}

#[cfg(target_os = "windows")]
fn suppress_overlay_lighting(_verbose: bool) {
    prepare_unfocused_lighting(false);
}

#[cfg(target_os = "windows")]
fn set_controlled_by_foreground_app(enabled: bool) {
    let value = if enabled { 1 } else { 0 };
    let _ = reg_add_dword(r"HKCU\Software\Microsoft\Lighting", "ControlledByForegroundApp", value);
    for device_key in reg_query_subkeys(r"HKCU\Software\Microsoft\Lighting\Devices") {
        let _ = reg_add_dword(&device_key, "ControlledByForegroundApp", value);
    }
}

#[cfg(target_os = "windows")]
fn windows_lighting_color_dword(r: u8, g: u8, b: u8) -> u32 {
    0xFF00_0000 | ((b as u32) << 16) | ((g as u32) << 8) | r as u32
}

#[cfg(target_os = "windows")]
fn zone_color(rgb: &[u8; 12], zone: usize) -> (u8, u8, u8) {
    (rgb[zone * 3], rgb[zone * 3 + 1], rgb[zone * 3 + 2])
}

#[cfg(target_os = "windows")]
fn zone_color_distance(a: (u8, u8, u8), b: (u8, u8, u8)) -> u16 {
    a.0.abs_diff(b.0) as u16 + a.1.abs_diff(b.1) as u16 + a.2.abs_diff(b.2) as u16
}

/// Map 4 keyboard zones onto Windows hold lighting (solid or left→right gradient).
#[cfg(target_os = "windows")]
fn pick_hold_zone_colors(rgb: &[u8; 12]) -> Option<((u8, u8, u8), (u8, u8, u8), u32)> {
    let zones = [
        zone_color(rgb, 0),
        zone_color(rgb, 1),
        zone_color(rgb, 2),
        zone_color(rgb, 3),
    ];
    if zones.iter().all(|c| c.0.max(c.1).max(c.2) < 24) {
        return None;
    }

    let mut max_dist = 0u16;
    for i in 0..4 {
        for j in (i + 1)..4 {
            max_dist = max_dist.max(zone_color_distance(zones[i], zones[j]));
        }
    }

    // Near-identical zones were being turned into a 2-color gradient because a
    // color-picker fade made them differ by a few counts.
    if max_dist < 30 {
        let color = zones
            .into_iter()
            .max_by_key(|c| c.0 as u16 + c.1 as u16 + c.2 as u16)
            .unwrap_or(zones[0]);
        return Some((color, color, 0));
    }

    Some((zones[0], zones[3], 6))
}

/// Copy live colors into Windows Dynamic Lighting so Windows can paint them
/// after this process is no longer the foreground lighting app.
#[cfg(target_os = "windows")]
fn sync_windows_hold_color(rgb: &[u8; 12], brightness: u8, log_handoff: bool) {
    let Some(((r0, g0, b0), (r1, g1, b1), effect)) = pick_hold_zone_colors(rgb) else {
        return;
    };
    let color = windows_lighting_color_dword(r0, g0, b0);
    let color2 = windows_lighting_color_dword(r1, g1, b1);
    let brightness = brightness.clamp(1, 100) as u32;

    static LAST: std::sync::Mutex<(u32, u32, u32, u32)> = std::sync::Mutex::new((0, 0, 0, 0));
    let colors_unchanged = if let Ok(mut last) = LAST.lock() {
        let same = *last == (color, color2, brightness, effect);
        *last = (color, color2, brightness, effect);
        same
    } else {
        false
    };

    if colors_unchanged {
        return;
    }

    if log_handoff {
        log_to_file(&format!(
            "WDL: handing colors to Windows hold lighting zones=[{},{},{} {},{},{} {},{},{} {},{},{}] -> r={} g={} b={} r2={} g2={} b2={} effect={}",
            rgb[0], rgb[1], rgb[2], rgb[3], rgb[4], rgb[5], rgb[6], rgb[7], rgb[8], rgb[9], rgb[10], rgb[11],
            r0, g0, b0, r1, g1, b1, effect
        ));
    }

    // Write the hold color first, but do not enable Ambient yet. Ambient + an
    // exclusive LampArray at the same time is what made the keys flicker.
    write_windows_hold_color_values(color, color2, effect, brightness);
}

#[cfg(target_os = "windows")]
fn write_windows_hold_color_values(color: u32, color2: u32, effect: u32, brightness: u32) {
    let root = r"Software\Microsoft\Lighting";
    write_hold_dwords(root, color, color2, effect, brightness);
    for device_key in lighting_device_subkeys() {
        write_hold_dwords(&device_key, color, color2, effect, brightness);
    }
}

#[cfg(target_os = "windows")]
fn write_hold_dwords(subkey: &str, color: u32, color2: u32, effect: u32, brightness: u32) {
    let _ = native_reg_dword(subkey, "EffectType", effect);
    let _ = native_reg_dword(subkey, "EffectMode", 0);
    let _ = native_reg_dword(subkey, "UseSystemAccentColor", 0);
    let _ = native_reg_dword(subkey, "Color", color);
    let _ = native_reg_dword(subkey, "Color2", color2);
    let _ = native_reg_dword(subkey, "Brightness", brightness);
}

#[cfg(target_os = "windows")]
fn lighting_device_subkeys() -> Vec<String> {
    static CACHE: std::sync::Mutex<Vec<String>> = std::sync::Mutex::new(Vec::new());
    if let Ok(cache) = CACHE.lock() {
        if !cache.is_empty() {
            return cache.clone();
        }
    }
    let keys = reg_query_subkeys(r"HKCU\Software\Microsoft\Lighting\Devices")
        .into_iter()
        .filter_map(|key| hkcu_subkey(&key).map(str::to_string))
        .collect::<Vec<_>>();
    if let Ok(mut cache) = CACHE.lock() {
        *cache = keys.clone();
    }
    keys
}

#[cfg(target_os = "windows")]
fn hkcu_subkey(key: &str) -> Option<&str> {
    let key = key.trim();
    key.strip_prefix("HKEY_CURRENT_USER\\")
        .or_else(|| key.strip_prefix("HKCU\\"))
}

#[cfg(target_os = "windows")]
fn native_reg_dword(subkey: &str, value: &str, data: u32) -> bool {
    use windows::Win32::Foundation::ERROR_SUCCESS;
    use windows::Win32::System::Registry::{
        RegSetKeyValueW, HKEY_CURRENT_USER, REG_DWORD,
    };

    let mut subkey_w: Vec<u16> = subkey.encode_utf16().collect();
    subkey_w.push(0);
    let mut value_w: Vec<u16> = value.encode_utf16().collect();
    value_w.push(0);
    let bytes = data.to_le_bytes();
    unsafe {
        RegSetKeyValueW(
            HKEY_CURRENT_USER,
            windows::core::PCWSTR(subkey_w.as_ptr()),
            windows::core::PCWSTR(value_w.as_ptr()),
            REG_DWORD.0,
            Some(bytes.as_ptr().cast()),
            bytes.len() as u32,
        ) == ERROR_SUCCESS
    }
}

#[cfg(target_os = "windows")]
fn enable_windows_hold_lighting() {
    let _ = reg_add_dword(r"HKCU\Software\Microsoft\Lighting", "AmbientLightingEnabled", 1);
    restore_windows_lighting_provider(r"HKCU\Software\Microsoft\Lighting\Providers", false);
    remove_legion_lighting_providers(r"HKCU\Software\Microsoft\Lighting\Providers");
    for device_key in reg_query_subkeys(r"HKCU\Software\Microsoft\Lighting\Devices") {
        let _ = reg_add_dword(&device_key, "AmbientLightingEnabled", 1);
        restore_windows_lighting_provider(&format!("{}\\Providers", device_key), false);
        remove_legion_lighting_providers(&format!("{}\\Providers", device_key));
    }
}

#[cfg(target_os = "windows")]
fn reg_add_dword(key: &str, value: &str, data: u32) -> bool {
    use std::os::windows::process::CommandExt;
    const CREATE_NO_WINDOW: u32 = 0x08000000;
    std::process::Command::new("reg")
        .args([
            "add",
            key,
            "/v",
            value,
            "/t",
            "REG_DWORD",
            "/d",
            &data.to_string(),
            "/f",
        ])
        .creation_flags(CREATE_NO_WINDOW)
        .output()
        .map(|output| output.status.success())
        .unwrap_or(false)
}

#[cfg(target_os = "windows")]
fn reg_add_sz(key: &str, value: &str, data: &str) -> bool {
    use std::os::windows::process::CommandExt;
    const CREATE_NO_WINDOW: u32 = 0x08000000;
    std::process::Command::new("reg")
        .args(["add", key, "/v", value, "/t", "REG_SZ", "/d", data, "/f"])
        .creation_flags(CREATE_NO_WINDOW)
        .output()
        .map(|output| output.status.success())
        .unwrap_or(false)
}

#[cfg(target_os = "windows")]
fn restore_windows_lighting_provider(key: &str, verbose: bool) {
    if reg_add_sz(key, "WindowsLighting", "1") && verbose {
        log_to_file(&format!("WDL: restored WindowsLighting provider in {}", key));
    }
}

#[cfg(target_os = "windows")]
fn disable_legion_lighting_background_apps(verbose: bool) {
    let root = r"HKCU\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications";
    for key in reg_query_subkeys(root) {
        if !key.to_ascii_lowercase().contains("legionlighting") {
            continue;
        }
        let disabled = reg_add_dword(&key, "Disabled", 1) && reg_add_dword(&key, "DisabledByUser", 1);
        if disabled && verbose {
            log_to_file(&format!("WDL: disabled Legion lighting background app ({})", key));
        }
    }
}

#[cfg(target_os = "windows")]
fn reg_query_subkeys(key: &str) -> Vec<String> {
    use std::os::windows::process::CommandExt;
    const CREATE_NO_WINDOW: u32 = 0x08000000;
    let output = std::process::Command::new("reg")
        .args(["query", key])
        .creation_flags(CREATE_NO_WINDOW)
        .output();
    let Ok(output) = output else {
        return Vec::new();
    };
    String::from_utf8_lossy(&output.stdout)
        .lines()
        .map(str::trim)
        .filter(|line| line.starts_with("HKEY_") && !line.eq_ignore_ascii_case(key))
        .map(|line| line.to_string())
        .collect()
}

#[cfg(target_os = "windows")]
fn remove_legion_lighting_providers(key: &str) {
    use std::os::windows::process::CommandExt;
    const CREATE_NO_WINDOW: u32 = 0x08000000;
    let output = std::process::Command::new("reg")
        .args(["query", key])
        .creation_flags(CREATE_NO_WINDOW)
        .output();
    let Ok(output) = output else {
        return;
    };
    for line in String::from_utf8_lossy(&output.stdout).lines() {
        let line = line.trim();
        let lower = line.to_ascii_lowercase();
        if !lower.contains("legionlighting") {
            continue;
        }
        let Some(value_name) = line.split_whitespace().next() else {
            continue;
        };
        let deleted = std::process::Command::new("reg")
            .args(["delete", key, "/v", value_name, "/f"])
            .creation_flags(CREATE_NO_WINDOW)
            .output()
            .map(|output| output.status.success())
            .unwrap_or(false);
        if deleted {
            log_to_file(&format!(
                "WDL: removed Legion overlay provider {} from {}",
                value_name, key
            ));
        }
    }
}

#[cfg(target_os = "windows")]
fn reg_query_value(key: &str, value: &str) -> String {
    use std::os::windows::process::CommandExt;
    const CREATE_NO_WINDOW: u32 = 0x08000000;
    let output = std::process::Command::new("reg")
        .args(["query", key, "/v", value])
        .creation_flags(CREATE_NO_WINDOW)
        .output();
    let Ok(output) = output else {
        return "?".into();
    };
    let text = String::from_utf8_lossy(&output.stdout);
    for line in text.lines() {
        let line = line.trim();
        if line.to_ascii_lowercase().contains(&value.to_ascii_lowercase()) {
            return line.split_whitespace().last().unwrap_or("?").to_string();
        }
    }
    "missing".into()
}

#[cfg(target_os = "windows")]
fn reg_query_providers(key: &str) -> String {
    use std::os::windows::process::CommandExt;
    const CREATE_NO_WINDOW: u32 = 0x08000000;
    let output = std::process::Command::new("reg")
        .args(["query", key])
        .creation_flags(CREATE_NO_WINDOW)
        .output();
    let Ok(output) = output else {
        return "?".into();
    };
    let mut parts = Vec::new();
    for line in String::from_utf8_lossy(&output.stdout).lines() {
        let line = line.trim();
        if line.is_empty() || line.starts_with("HKEY_") {
            continue;
        }
        let lower = line.to_ascii_lowercase();
        if lower.contains("windowslighting") || lower.contains("legionlighting") {
            parts.push(line.split_whitespace().take(3).collect::<Vec<_>>().join("="));
        }
    }
    if parts.is_empty() {
        "none".into()
    } else {
        parts.join("; ")
    }
}

/// Snapshot every lighting writer that can fight over the same keys.
#[cfg(target_os = "windows")]
fn log_fight_snapshot(tag: &str, holding_lamp: bool, window_active: bool, static_frame: bool) {
    let root = r"HKCU\Software\Microsoft\Lighting";
    let ambient = reg_query_value(root, "AmbientLightingEnabled");
    let fg = reg_query_value(root, "ControlledByForegroundApp");
    let effect = reg_query_value(root, "EffectType");
    let color = reg_query_value(root, "Color");
    let providers = reg_query_providers(r"HKCU\Software\Microsoft\Lighting\Providers");
    let fighting = holding_lamp && ambient.eq_ignore_ascii_case("0x1")
        || providers.to_ascii_lowercase().contains("legionlighting");
    log_to_file(&format!(
        "FIGHT: {} holding_lamp={} window_active={} static_frame={} ambient={} fg_app={} effect={} color={} providers=[{}] clash={}",
        tag, holding_lamp, window_active, static_frame, ambient, fg, effect, color, providers, fighting
    ));
}

#[cfg(target_os = "windows")]
fn stop_legion_lighting_helpers() {
    use std::os::windows::process::CommandExt;
    const CREATE_NO_WINDOW: u32 = 0x08000000;
    // Lighting helpers only — do not kill the main Legion Space / Vantage UI.
    for image in [
        "LegionLightingController.exe",
        "LightingService.exe",
        "LenovoLighting.exe",
        "LenovoLightingService.exe",
    ] {
        let output = std::process::Command::new("taskkill")
            .args(["/IM", image, "/F"])
            .creation_flags(CREATE_NO_WINDOW)
            .output();
        if let Ok(output) = output {
            if output.status.success() {
                log_to_file(&format!("WDL: stopped overlay process {}", image));
            }
        }
    }
}

#[cfg(target_os = "windows")]
use windows::{
    Devices::{
        Enumeration::DeviceInformation,
        Lights::{
            LampArray,
            Effects::{
                LampArrayCustomEffect,
                LampArrayEffectPlaylist,
                LampArrayRepetitionMode,
                LampArrayUpdateRequestedEventArgs,
            },
        },
    },
    Foundation::{TimeSpan, TypedEventHandler},
    UI::Color,
    core::HSTRING,
};

const KNOWN_DEVICE_INFOS: [(u16, u16, u16, u16); 12] = [
    (0x048d, 0xc995, 0xff89, 0x00cc), // 2024 Pro
    (0x048d, 0xc994, 0xff89, 0x00cc), // 2024
    (0x048d, 0xc993, 0xff89, 0x00cc), // 2024 LOQ
    (0x048d, 0xc985, 0xff89, 0x00cc), // 2023 Pro
    (0x048d, 0xc984, 0xff89, 0x00cc), // 2023
    (0x048d, 0xc983, 0xff89, 0x00cc), // 2023 LOQ
    (0x048d, 0xc975, 0xff89, 0x00cc), // 2022
    (0x048d, 0xc973, 0xff89, 0x00cc), // 2022 Ideapad
    (0x048d, 0xc965, 0xff89, 0x00cc), // 2021
    (0x048d, 0xc963, 0xff89, 0x00cc), // 2021 Ideapad
    (0x048d, 0xc955, 0xff89, 0x00cc), // 2020
    (0x048d, 0xc693, 0xff89, 0x00cc), // 2025 LOQ 15IRX10
];

// --- HID LampArray (Windows Dynamic Lighting) constants ---
const LAMP_ARRAY_USAGE_PAGE: u16 = 0x0059;
const LAMP_ARRAY_USAGE: u16 = 0x0001;

// HID Lighting and Illumination Usage IDs (HID Usage Tables §19)
const USAGE_LAMP_ARRAY_ATTRIBUTES_REPORT: u32 = 0x02;
const USAGE_LAMP_ATTRIBUTES_REQUEST_REPORT: u32 = 0x20;
const USAGE_LAMP_ATTRIBUTES_RESPONSE_REPORT: u32 = 0x22;
const USAGE_LAMP_MULTI_UPDATE_REPORT: u32 = 0x50;
const USAGE_LAMP_RANGE_UPDATE_REPORT: u32 = 0x60;
const USAGE_LAMP_ARRAY_CONTROL_REPORT: u32 = 0x70;

/// LampArray HID report IDs, discovered from the device's report descriptor.
#[derive(Clone, Copy, Debug)]
struct LampArrayReportIds {
    attributes: u8,
    lamp_attr_request: u8,
    lamp_attr_response: u8,
    multi_update: u8,
    range_update: u8,
    control: u8,
}

/// Tracks which HID write method works for this device, so we don't
/// retry all methods on every refresh call.
#[derive(Clone, Copy, Debug)]
enum WriteMethod {
    /// Not yet determined — will probe on first refresh.
    Unknown,
    /// `send_feature_report` (works for most 2020-2024 models).
    FeatureReport,
    /// `write` (works for newer models like LOQ 15IRX10).
    Write,
    /// `write` with a 0x00 report-ID byte prepended.
    WriteWithReportId,
}

#[cfg(target_os = "windows")]
#[derive(Clone, Debug)]
enum VendorWriteMode {
    Probe,
    Disabled,
    Write,
    WriteWithZeroReportId,
    FeatureReport,
    FeatureWithZeroReportId,
    WriteWithReportId(u8),
    FeatureWithReportId(u8),
}

#[cfg(target_os = "windows")]
#[derive(Clone, Debug)]
struct Gen7LedGroup {
    mode: u8,
    speed: u8,
    spin: u8,
    direction: u8,
    color_mode: u8,
    colors: Vec<[u8; 3]>,
    leds: Vec<u16>,
}

#[cfg(target_os = "windows")]
#[derive(Clone, Debug, PartialEq, Eq)]
struct Gen7ApplyState {
    rgb_values: [u8; 12],
    brightness: u8,
    speed: u8,
    effect_type: BaseEffects,
}

#[cfg(target_os = "windows")]
#[derive(Clone, Debug)]
struct Gen7ControllerState {
    profile_id: u8,
    groups: Vec<Gen7LedGroup>,
    #[allow(dead_code)]
    direct_leds: Vec<u16>,
    direct_enabled: bool,
    last_applied: Option<Gen7ApplyState>,
}

#[cfg(target_os = "windows")]
struct VendorHidTarget {
    label: String,
    product_id: u16,
    usage_page: u16,
    usage: u16,
    device: HidDevice,
    output_report_ids: Vec<u8>,
    feature_report_ids: Vec<u8>,
    feature_report_lengths: BTreeMap<u8, usize>,
    gen7_controller: Option<Gen7ControllerState>,
    mode: VendorWriteMode,
}

#[cfg(target_os = "windows")]
fn prefixed_payload(report_id: u8, payload: &[u8; 33]) -> [u8; 34] {
    let mut buf = [0u8; 34];
    buf[0] = report_id;
    buf[1..].copy_from_slice(payload);
    buf
}

#[cfg(target_os = "windows")]
const GEN7_PACKET_SIZE: usize = 960;

#[cfg(target_os = "windows")]
const GEN7_REPORT_ID: u8 = 0x07;

#[cfg(target_os = "windows")]
const GEN7_GET_ACTIVE_PROFILE: u8 = 0xca;

#[cfg(target_os = "windows")]
const GEN7_DIRECT_MODE: u8 = 0xa1;

#[cfg(target_os = "windows")]
const GEN7_SAVE_PROFILE: u8 = 0xcb;

#[cfg(target_os = "windows")]
const GEN7_GET_PROFILE: u8 = 0xcc;

#[cfg(target_os = "windows")]
const GEN7_GET_BRIGHTNESS: u8 = 0xcd;

#[cfg(target_os = "windows")]
const GEN7_SET_BRIGHTNESS: u8 = 0xce;

#[cfg(target_os = "windows")]
const GEN7_SET_DIRECT_MODE: u8 = 0xd0;

#[cfg(target_os = "windows")]
const GEN7_MODE_COLOR_PULSE: u8 = 0x04;

#[cfg(target_os = "windows")]
const GEN7_MODE_COLOR_WAVE: u8 = 0x05;

#[cfg(target_os = "windows")]
const GEN7_MODE_SMOOTH: u8 = 0x06;

#[cfg(target_os = "windows")]
const GEN7_MODE_STATIC: u8 = 0x0b;

#[cfg(target_os = "windows")]
fn set_gen7_payload_length(product_id: u16, buffer: &mut [u8; GEN7_PACKET_SIZE], payload_length: u16) {
    if product_id != 0xc197 {
        return;
    }

    buffer[2] = (payload_length & 0xff) as u8;
    buffer[3] = (payload_length >> 8) as u8;
}

#[cfg(target_os = "windows")]
fn send_gen7_feature_report(
    device: &HidDevice,
    buffer: &[u8; GEN7_PACKET_SIZE],
    context: &str,
) -> std::result::Result<(), hidapi::HidError> {
    if reverse_engineering_mode_enabled() {
        let len_hint = if buffer[1] == GEN7_SAVE_PROFILE {
            buffer[2] as usize
        } else {
            8
        }
        .clamp(0, GEN7_PACKET_SIZE);
        let preview = buffer[..len_hint.max(6)]
            .iter()
            .take(48)
            .map(|b| format!("{:02x}", b))
            .collect::<Vec<_>>()
            .join(" ");
        log_to_file(&format!(
            "RE-GEN7 [{}]: send len_hint={} bytes={}",
            context, len_hint, preview
        ));
    }

    let _ = device.send_feature_report(buffer)?;
    Ok(())
}

#[cfg(target_os = "windows")]
fn request_gen7_feature_report(
    device: &HidDevice,
    request: &mut [u8; GEN7_PACKET_SIZE],
    context: &str,
) -> std::result::Result<Vec<u8>, hidapi::HidError> {
    let _ = device.send_feature_report(request)?;

    let mut response = [0u8; GEN7_PACKET_SIZE];
    response[0] = GEN7_REPORT_ID;
    let len = device.get_feature_report(&mut response)?;
    let data = response[..len].to_vec();

    if reverse_engineering_mode_enabled() {
        let preview = data
            .iter()
            .take(64)
            .map(|b| format!("{:02x}", b))
            .collect::<Vec<_>>()
            .join(" ");
        log_to_file(&format!(
            "RE-GEN7 [{}]: recv len={} bytes={}{}",
            context,
            data.len(),
            preview,
            if data.len() > 64 { " ..." } else { "" }
        ));
    }

    Ok(data)
}

#[cfg(target_os = "windows")]
fn parse_gen7_profile_groups(response: &[u8]) -> Vec<Gen7LedGroup> {
    let mut groups = Vec::new();
    let mut index = 7usize;

    while index < response.len() && response[index] != 0x00 {
        index += 1;
        if index >= response.len() {
            break;
        }

        let mut group = Gen7LedGroup {
            mode: GEN7_MODE_STATIC,
            speed: 1,
            spin: 0,
            direction: 0,
            color_mode: 0,
            colors: Vec::new(),
            leds: Vec::new(),
        };

        let setting_count = response[index] as usize;
        index += 1;
        for _ in 0..setting_count {
            if index + 1 >= response.len() {
                return groups;
            }
            match response[index] {
                0x01 => group.mode = response[index + 1],
                0x02 => group.speed = response[index + 1],
                0x03 => group.spin = response[index + 1],
                0x04 => group.direction = response[index + 1],
                0x05 => group.color_mode = response[index + 1],
                _ => {}
            }
            index += 2;
        }

        if index >= response.len() {
            break;
        }

        let color_count = response[index] as usize;
        index += 1;
        for _ in 0..color_count {
            if index + 2 >= response.len() {
                return groups;
            }
            group.colors.push([response[index], response[index + 1], response[index + 2]]);
            index += 3;
        }

        if index >= response.len() {
            break;
        }

        let led_count = response[index] as usize;
        index += 1;
        for _ in 0..led_count {
            if index + 1 >= response.len() {
                return groups;
            }
            group.leds.push(u16::from_le_bytes([response[index], response[index + 1]]));
            index += 2;
        }

        groups.push(group);
    }

    groups
}

#[cfg(target_os = "windows")]
fn collect_gen7_direct_leds(groups: &[Gen7LedGroup]) -> Vec<u16> {
    let mut leds = Vec::new();
    for group in groups {
        for led in &group.leds {
            if !leds.contains(led) {
                leds.push(*led);
            }
        }
    }
    leds
}

#[cfg(target_os = "windows")]
fn send_gen7_direct_mode(
    target: &VendorHidTarget,
    profile_id: u8,
    enabled: bool,
) -> std::result::Result<(), hidapi::HidError> {
    let mut packet = [0u8; GEN7_PACKET_SIZE];
    packet[0] = GEN7_REPORT_ID;
    packet[1] = GEN7_SET_DIRECT_MODE;
    packet[2] = 0xc0;
    packet[3] = 0x03;
    packet[4] = if enabled { 0x01 } else { 0x02 };
    packet[5] = profile_id;
    set_gen7_payload_length(target.product_id, &mut packet, 2);
    send_gen7_feature_report(
        &target.device,
        &packet,
        &format!("{} direct_mode_{}", target.label, if enabled { "on" } else { "off" }),
    )
}

#[cfg(target_os = "windows")]
#[allow(dead_code)]
fn build_gen7_direct_colors(led_count: usize, rgb_values: [u8; 12]) -> Vec<[u8; 3]> {
    let palette = build_gen7_palette(rgb_values);
    if palette.is_empty() || led_count == 0 {
        return Vec::new();
    }

    (0..led_count)
        .map(|index| {
            let zone = ((index * palette.len()) / led_count).min(palette.len() - 1);
            palette[zone]
        })
        .collect()
}

#[cfg(target_os = "windows")]
#[allow(dead_code)]
fn build_gen7_direct_packet(
    product_id: u16,
    leds: &[u16],
    colors: &[[u8; 3]],
) -> [u8; GEN7_PACKET_SIZE] {
    let mut buffer = [0u8; GEN7_PACKET_SIZE];
    let mut index = 0usize;

    buffer[index] = GEN7_REPORT_ID;
    index += 1;
    buffer[index] = GEN7_DIRECT_MODE;
    index += 1;

    if product_id == 0xc197 {
        buffer[index] = 0x00;
        index += 1;
        buffer[index] = 0x00;
        index += 1;

        let mut count = 0usize;
        for (led, color) in leds.iter().zip(colors.iter()) {
            if index + 5 > GEN7_PACKET_SIZE {
                break;
            }
            let [low, high] = led.to_le_bytes();
            buffer[index] = low;
            buffer[index + 1] = high;
            buffer[index + 2] = color[0];
            buffer[index + 3] = color[1];
            buffer[index + 4] = color[2];
            index += 5;
            count += 1;
        }

        set_gen7_payload_length(product_id, &mut buffer, (count * 5) as u16);
        return buffer;
    }

    buffer[index] = 0xc0;
    index += 1;
    buffer[index] = 0x03;
    index += 1;
    buffer[index] = 0x03;
    index += 1;

    for (led, color) in leds.iter().zip(colors.iter()) {
        if index + 5 > GEN7_PACKET_SIZE {
            break;
        }
        let [low, high] = led.to_le_bytes();
        buffer[index] = low;
        buffer[index + 1] = high;
        buffer[index + 2] = color[0];
        buffer[index + 3] = color[1];
        buffer[index + 4] = color[2];
        index += 5;
    }

    buffer
}

#[cfg(target_os = "windows")]
fn default_loq_gen7_groups() -> Vec<Gen7LedGroup> {
    (0..4)
        .map(|zone| {
            let start = zone * 6;
            Gen7LedGroup {
                mode: GEN7_MODE_STATIC,
                speed: 1,
                spin: 0,
                direction: 0,
                color_mode: 0x02,
                colors: vec![[0, 0, 0]],
                leds: (start..start + 6).map(|id| id as u16).collect(),
            }
        })
        .collect()
}

#[cfg(target_os = "windows")]
fn fallback_loq_gen7_controller(product_id: u16) -> Option<Gen7ControllerState> {
    if product_id != 0xc693 {
        return None;
    }
    Some(Gen7ControllerState {
        profile_id: 1,
        groups: default_loq_gen7_groups(),
        direct_leds: (0..24).collect(),
        direct_enabled: false,
        last_applied: None,
    })
}

#[cfg(target_os = "windows")]
fn try_load_gen7_controller(
    device: &HidDevice,
    label: &str,
    product_id: u16,
) -> Option<Gen7ControllerState> {
    let mut profile_request = [0u8; GEN7_PACKET_SIZE];
    profile_request[0] = GEN7_REPORT_ID;
    profile_request[1] = GEN7_GET_ACTIVE_PROFILE;
    profile_request[2] = 0xc0;
    profile_request[3] = 0x03;
    set_gen7_payload_length(product_id, &mut profile_request, 1);
    let profile_response = match request_gen7_feature_report(
        device,
        &mut profile_request,
        &format!("{} active_profile", label),
    ) {
        Ok(response) => response,
        Err(e) => {
            log_to_file(&format!(
                "RE-GEN7 [{}]: active profile query failed: {}",
                label, e
            ));
            if let Some(controller) = fallback_loq_gen7_controller(product_id) {
                log_to_file(&format!(
                    "RE-GEN7 [{}]: using built-in LOQ profile fallback after active-profile failure",
                    label
                ));
                return Some(controller);
            }
            return None;
        }
    };
    let profile_id = profile_response.get(4).copied().unwrap_or(1);

    let mut brightness_request = [0u8; GEN7_PACKET_SIZE];
    brightness_request[0] = GEN7_REPORT_ID;
    brightness_request[1] = GEN7_GET_BRIGHTNESS;
    brightness_request[2] = 0xc0;
    brightness_request[3] = 0x03;
    set_gen7_payload_length(product_id, &mut brightness_request, 1);
    let _ = request_gen7_feature_report(
        device,
        &mut brightness_request,
        &format!("{} brightness", label),
    );

    let mut group_request = [0u8; GEN7_PACKET_SIZE];
    group_request[0] = GEN7_REPORT_ID;
    group_request[1] = GEN7_GET_PROFILE;
    group_request[2] = 0xc0;
    group_request[3] = 0x03;
    group_request[4] = profile_id;
    set_gen7_payload_length(product_id, &mut group_request, (GEN7_PACKET_SIZE - 4) as u16);
    let response = match request_gen7_feature_report(
        device,
        &mut group_request,
        &format!("{} profile_{}", label, profile_id),
    ) {
        Ok(response) => response,
        Err(e) => {
            log_to_file(&format!(
                "RE-GEN7 [{}]: profile query failed: {}",
                label, e
            ));
            if let Some(mut controller) = fallback_loq_gen7_controller(product_id) {
                controller.profile_id = profile_id;
                log_to_file(&format!(
                    "RE-GEN7 [{}]: using built-in LOQ profile fallback after profile query failure",
                    label
                ));
                return Some(controller);
            }
            return None;
        }
    };

    let groups = parse_gen7_profile_groups(&response);
    let groups = if groups.is_empty() {
        log_to_file(&format!(
            "RE-GEN7 [{}]: profile {} returned no groups, using built-in 4-zone map",
            label, profile_id
        ));
        default_loq_gen7_groups()
    } else {
        groups
    };

    let direct_leds = collect_gen7_direct_leds(&groups);
    log_to_file(&format!(
        "RE-GEN7 [{}]: active profile={} parsed_groups={} direct_leds={}",
        label,
        profile_id,
        groups.len(),
        direct_leds.len()
    ));

    Some(Gen7ControllerState {
        profile_id,
        groups,
        direct_leds,
        direct_enabled: false,
        last_applied: None,
    })
}

#[cfg(target_os = "windows")]
fn build_gen7_palette(rgb_values: [u8; 12]) -> Vec<[u8; 3]> {
    rgb_values
        .chunks_exact(3)
        .map(|chunk| [chunk[0], chunk[1], chunk[2]])
        .collect()
}

#[cfg(target_os = "windows")]
fn map_gen7_mode(effect_type: BaseEffects) -> u8 {
    match effect_type {
        BaseEffects::Static => GEN7_MODE_STATIC,
        BaseEffects::Breath => GEN7_MODE_COLOR_PULSE,
        BaseEffects::Smooth => GEN7_MODE_SMOOTH,
        BaseEffects::LeftWave | BaseEffects::RightWave => GEN7_MODE_COLOR_WAVE,
    }
}

#[cfg(target_os = "windows")]
fn map_gen7_direction(effect_type: BaseEffects) -> u8 {
    match effect_type {
        BaseEffects::LeftWave => 0x03,
        BaseEffects::RightWave => 0x04,
        _ => 0x00,
    }
}

#[cfg(target_os = "windows")]
fn map_gen7_speed(speed: u8) -> u8 {
    speed.clamp(1, 3)
}

#[cfg(target_os = "windows")]
fn map_gen7_brightness(brightness: u8) -> u8 {
    (((brightness.clamp(1, 100) - 1) as u16 * 9) / 99) as u8
}

#[cfg(target_os = "windows")]
fn build_gen7_profile_packet(
    product_id: u16,
    profile_id: u8,
    groups: &[Gen7LedGroup],
) -> [u8; GEN7_PACKET_SIZE] {
    let mut buffer = [0u8; GEN7_PACKET_SIZE];
    let mut index = 0usize;

    buffer[index] = GEN7_REPORT_ID;
    index += 1;
    buffer[index] = GEN7_SAVE_PROFILE;
    index += 1;
    buffer[index] = 0xc0;
    index += 1;
    buffer[index] = 0x03;
    index += 1;
    buffer[index] = profile_id;
    index += 1;
    buffer[index] = 0x01;
    index += 1;
    buffer[index] = 0x01;
    index += 1;

    for (group_index, group) in groups.iter().enumerate() {
        if index + 16 >= GEN7_PACKET_SIZE {
            break;
        }

        buffer[index] = (group_index + 1) as u8;
        index += 1;
        buffer[index] = 0x06;
        index += 1;

        let settings = [
            (0x01, group.mode),
            (0x02, group.speed),
            (0x03, group.spin),
            (0x04, group.direction),
            (0x05, group.color_mode),
            (0x06, 0x00),
        ];
        for (key, value) in settings {
            buffer[index] = key;
            buffer[index + 1] = value;
            index += 2;
        }

        buffer[index] = group.colors.len().min(u8::MAX as usize) as u8;
        index += 1;
        for color in &group.colors {
            if index + 3 > GEN7_PACKET_SIZE {
                break;
            }
            buffer[index] = color[0];
            buffer[index + 1] = color[1];
            buffer[index + 2] = color[2];
            index += 3;
        }

        let led_count = group
            .leds
            .len()
            .min(((GEN7_PACKET_SIZE - index) / 2).min(u8::MAX as usize));
        buffer[index] = led_count as u8;
        index += 1;
        for led in group.leds.iter().take(led_count) {
            let [low, high] = led.to_le_bytes();
            buffer[index] = low;
            buffer[index + 1] = high;
            index += 2;
        }
    }

    if product_id == 0xc197 {
        set_gen7_payload_length(product_id, &mut buffer, (index.saturating_sub(4)) as u16);
    } else {
        buffer[2] = index.min(u8::MAX as usize) as u8;
    }

    buffer
}

#[cfg(target_os = "windows")]
fn apply_gen7_controller(
    target: &mut VendorHidTarget,
    controller: &mut Gen7ControllerState,
    rgb_values: [u8; 12],
    brightness: u8,
    speed: u8,
    effect_type: BaseEffects,
    tick: u64,
) {
    let _ = tick;
    let apply_state = Gen7ApplyState {
        rgb_values,
        brightness,
        speed,
        effect_type,
    };
    if controller.last_applied.as_ref() == Some(&apply_state) {
        return;
    }
    if !is_vivid_rgb_frame(&rgb_values) {
        return;
    }

    let palette = build_gen7_palette(rgb_values);
    if palette.is_empty() {
        return;
    }

    let mode = map_gen7_mode(effect_type);
    let direction = map_gen7_direction(effect_type);
    let mapped_speed = map_gen7_speed(speed);
    let mapped_brightness = map_gen7_brightness(brightness);

    let mut brightness_packet = [0u8; GEN7_PACKET_SIZE];
    brightness_packet[0] = GEN7_REPORT_ID;
    brightness_packet[1] = GEN7_SET_BRIGHTNESS;
    brightness_packet[2] = 0xc0;
    brightness_packet[3] = 0x03;
    brightness_packet[4] = mapped_brightness;
    set_gen7_payload_length(target.product_id, &mut brightness_packet, 1);
    if let Err(e) = send_gen7_feature_report(
        &target.device,
        &brightness_packet,
        &format!("{} set_brightness", target.label),
    ) {
        log_to_file(&format!(
            "RE-GEN7 [{}]: set brightness failed: {}",
            target.label, e
        ));
        return;
    }

    // Direct/LED overlay sits *above* the keyboard's built-in FN+Space profile.
    // Legion Space uses that overlay; it vanishes on click-away. Save the onboard
    // profile instead so firmware keeps the colors.
    if controller.direct_enabled {
        if send_gen7_direct_mode(target, controller.profile_id, false).is_ok() {
            controller.direct_enabled = false;
            log_to_file(&format!(
                "RE-GEN7 [{}]: left overlay/direct mode, writing built-in profile",
                target.label
            ));
        }
    }

    let mut groups = if controller.groups.is_empty() {
        default_loq_gen7_groups()
    } else {
        controller.groups.clone()
    };

    // Firmware often stores one 24-lamp group. Split it so 4-zone colors
    // become the built-in profile instead of only the first zone.
    if matches!(effect_type, BaseEffects::Static)
        && groups.len() == 1
        && groups[0].leds.len() >= 8
        && palette.len() >= 2
    {
        let leds = groups[0].leds.clone();
        let chunk = (leds.len() + 3) / 4;
        groups = (0..4)
            .filter_map(|zone| {
                let start = zone * chunk;
                if start >= leds.len() {
                    return None;
                }
                let end = (start + chunk).min(leds.len());
                Some(Gen7LedGroup {
                    mode: GEN7_MODE_STATIC,
                    speed: mapped_speed,
                    spin: 0x00,
                    direction: direction,
                    color_mode: 0x02,
                    colors: vec![palette[zone.min(palette.len() - 1)]],
                    leds: leds[start..end].to_vec(),
                })
            })
            .collect();
    }
    for (group_index, group) in groups.iter_mut().enumerate() {
        group.mode = mode;
        group.speed = mapped_speed;
        group.spin = 0x00;
        group.direction = direction;
        group.color_mode = 0x02;
        group.colors = match effect_type {
            BaseEffects::Static => vec![palette[group_index.min(palette.len() - 1)]],
            _ => palette.clone(),
        };
    }

    let profile_packet = build_gen7_profile_packet(target.product_id, controller.profile_id, &groups);
    match send_gen7_feature_report(
        &target.device,
        &profile_packet,
        &format!("{} save_profile", target.label),
    ) {
        Ok(()) => {
            if controller.last_applied.is_none() {
                log_to_file(&format!(
                    "RE-GEN7 [{}]: saved built-in keyboard profile {} ({} groups)",
                    target.label,
                    controller.profile_id,
                    groups.len()
                ));
            }
            controller.last_applied = Some(apply_state);
            controller.groups = groups;
        }
        Err(e) => log_to_file(&format!(
            "RE-GEN7 [{}]: save profile failed: {}",
            target.label, e
        )),
    }
}

/// Which communication protocol the Keyboard uses.
#[derive(Clone, Copy, Debug)]
enum Protocol {
    /// Legacy Lenovo-specific 0xCC/0x16 payload protocol (2020-2024 models).
    Legacy(WriteMethod),
    /// Raw HID LampArray protocol (direct HID communication).
    LampArrayHid {
        report_ids: LampArrayReportIds,
        lamp_count: u16,
    },
    /// Raw HID LampArray with software-managed effects/brightness.
    /// Used for LOQ 15IRX10 so we avoid the WinRT path entirely.
    LampArrayHidManaged,
    /// Windows Dynamic Lighting WinRT API (most reliable for LOQ 15IRX10 etc.).
    #[cfg(target_os = "windows")]
    WindowsDynamicLighting {
        lamp_count: u16,
    },
}

/// Shared color state for the Windows Dynamic Lighting effect callback.
#[cfg(target_os = "windows")]
struct WinLampColors {
    rgb_values: [u8; 12],
    brightness: u8,
    speed: u8,
    effect_type: BaseEffects,
    lamp_count: u16,
}

pub const SPEED_RANGE: std::ops::RangeInclusive<u8> = 1..=4;
pub const BRIGHTNESS_RANGE: std::ops::RangeInclusive<u8> = 1..=2;
pub const ZONE_RANGE: std::ops::RangeInclusive<u8> = 0..=3;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BaseEffects {
    Static,
    Breath,
    Smooth,
    LeftWave,
    RightWave,
}


pub struct LightingState {
    effect_type: BaseEffects,
    speed: u8,
    brightness: u8,
    rgb_values: [u8; 12],
}

pub struct Keyboard {
    keyboard_hid: Option<HidDevice>,
    #[cfg(target_os = "windows")]
    win_lamp_array: Arc<std::sync::Mutex<Option<LampArray>>>,
    #[cfg(target_os = "windows")]
    win_color_state: Option<Arc<std::sync::Mutex<WinLampColors>>>,
    #[cfg(target_os = "windows")]
    _win_playlist: Option<LampArrayEffectPlaylist>,
    #[cfg(target_os = "windows")]
    _win_keepalive: Option<std::thread::JoinHandle<()>>,
    current_state: LightingState,
    stop_signal: Arc<AtomicBool>,
    window_active: Arc<AtomicBool>,
    protocol: Protocol,
}

#[allow(dead_code)]
impl Keyboard {
    pub fn window_active_handle(&self) -> Arc<AtomicBool> {
        self.window_active.clone()
    }

    /// Access the underlying HID device (panics if not available — only for HID protocols).
    fn hid_device(&self) -> &HidDevice {
        self.keyboard_hid.as_ref().expect("internal: HID device required for this protocol")
    }

    fn build_payload(&self) -> Result<[u8; 33]> {
        let keyboard_state = &self.current_state;

        if !SPEED_RANGE.contains(&keyboard_state.speed) {
            return Err(RangeError { kind: RangeErrorKind::Speed }.into());
        }
        if !BRIGHTNESS_RANGE.contains(&keyboard_state.brightness) {
            return Err(RangeError { kind: RangeErrorKind::Brightness }.into());
        }

        let mut payload: [u8; 33] = [0; 33];
        payload[0] = 0xcc;
        payload[1] = 0x16;
        payload[2] = match keyboard_state.effect_type {
            BaseEffects::Static => 0x01,
            BaseEffects::Breath => 0x03,
            BaseEffects::Smooth => 0x06,
            BaseEffects::LeftWave => {
                payload[19] = 0x1;
                0x04
            }
            BaseEffects::RightWave => {
                payload[18] = 0x1;
                0x04
            }
        };

        payload[3] = keyboard_state.speed;
        payload[4] = keyboard_state.brightness;

        if let BaseEffects::Static | BaseEffects::Breath = keyboard_state.effect_type {
            payload[5..(12 + 5)].copy_from_slice(&keyboard_state.rgb_values[..12]);
        };

        Ok(payload)
    }

    pub fn refresh(&mut self) -> Result<()> {
        match self.protocol {
            #[cfg(target_os = "windows")]
            Protocol::WindowsDynamicLighting { lamp_count } => {
                // Update shared color state; the effect callback applies it at ~20fps
                if let Some(ref shared) = self.win_color_state {
                    let mut state = shared.lock().unwrap();
                    state.rgb_values = self.current_state.rgb_values;
                    state.brightness = self.current_state.brightness;
                    state.speed = self.current_state.speed;
                    state.effect_type = self.current_state.effect_type;
                }

                // Direct write as backup for when effect playlist is paused (app backgrounded)
                if let Ok(guard) = self.win_lamp_array.lock() {
                    if let Some(ref lamp_array) = *guard {
                    let colors = &self.current_state.rgb_values;
                    let lc = lamp_count as usize;
                    let zones = std::cmp::min(4, lc);
                    let intensity = (self.current_state.brightness as f64 / 100.0).clamp(0.01, 1.0);

                    for z in 0..zones {
                        let zone_start = (z * lc) / zones;
                        let zone_end = ((z + 1) * lc) / zones;
                        if zone_start >= zone_end { continue; }
                        let r = (colors[z * 3] as f64 * intensity) as u8;
                        let g = (colors[z * 3 + 1] as f64 * intensity) as u8;
                        let b = (colors[z * 3 + 2] as f64 * intensity) as u8;
                        let color = Color { A: 255, R: r, G: g, B: b };
                        let indices: Vec<i32> = (zone_start..zone_end).map(|i| i as i32).collect();
                        let _ = lamp_array.SetSingleColorForIndices(color, &indices);
                    }
                    }
                }

                Ok(())
            }
            Protocol::LampArrayHid { report_ids, lamp_count } => {
                let device = self.hid_device();
                let colors = &self.current_state.rgb_values;
                let intensity = match self.current_state.brightness {
                    1 => 128u8,
                    _ => 255u8,
                };
                let zones = std::cmp::min(4, lamp_count as usize);

                // Build LampMultiUpdateReport (HID Usage 0x50)
                // Layout: [report_id, lamp_count, flags, lamp_ids×8 (u16 LE), colors×8 (RGBI)]
                let mut buf = [0u8; 51];
                buf[0] = report_ids.multi_update;
                buf[1] = zones as u8;
                buf[2] = 0x01; // flags: lampUpdateComplete

                // Lamp IDs (u16 LE, 8 slots starting at byte 3)
                for z in 0..zones {
                    buf[3 + z * 2] = z as u8;
                }
                // Colors (RGBI, 8 slots starting at byte 19)
                for z in 0..zones {
                    let co = 19 + z * 4;
                    let ro = z * 3;
                    buf[co] = colors[ro];           // Red
                    buf[co + 1] = colors[ro + 1];   // Green
                    buf[co + 2] = colors[ro + 2];   // Blue
                    buf[co + 3] = intensity;         // Intensity
                }

                #[cfg(debug_assertions)]
                eprintln!("[DEBUG] LampArray HID: sending MultiUpdate ({} zones, intensity={})", zones, intensity);

                device.write(&buf)?;
                Ok(())
            }
            Protocol::LampArrayHidManaged => {
                if let Some(ref shared) = self.win_color_state {
                    let mut state = shared.lock().unwrap();
                    state.rgb_values = self.current_state.rgb_values;
                    state.brightness = self.current_state.brightness;
                    state.speed = self.current_state.speed;
                    state.effect_type = self.current_state.effect_type;
                }
                // The keepalive thread handles all HID writes for this protocol.
                // refresh() just updates shared state here; writes happen in the keepalive loop.
                Ok(())
            }
            Protocol::Legacy(write_method) => {
                let payload = self.build_payload()?;

                #[cfg(debug_assertions)]
                eprintln!("[DEBUG] Sending payload (method: {:?}): {:02x?}", write_method, &payload[..5]);

                match write_method {
                    WriteMethod::FeatureReport => {
                        self.hid_device().send_feature_report(&payload)?;
                        return Ok(());
                    }
                    WriteMethod::Write => {
                        self.hid_device().write(&payload)?;
                        return Ok(());
                    }
                    WriteMethod::WriteWithReportId => {
                        let mut buf = [0u8; 34];
                        buf[0] = 0x00;
                        buf[1..].copy_from_slice(&payload);
                        self.hid_device().write(&buf)?;
                        return Ok(());
                    }
                    WriteMethod::Unknown => { /* probe below */ }
                }

                // Probe: try each method and remember which one works.
                if self.hid_device().send_feature_report(&payload).is_ok() {
                    #[cfg(debug_assertions)]
                    eprintln!("[DEBUG] Probe: feature report succeeded");
                    self.protocol = Protocol::Legacy(WriteMethod::FeatureReport);
                    return Ok(());
                }

                if self.hid_device().write(&payload).is_ok() {
                    #[cfg(debug_assertions)]
                    eprintln!("[DEBUG] Probe: write succeeded (LOQ 15IRX10 path)");
                    self.protocol = Protocol::Legacy(WriteMethod::Write);
                    return Ok(());
                }

                let mut buf = [0u8; 34];
                buf[0] = 0x00;
                buf[1..].copy_from_slice(&payload);
                if self.hid_device().write(&buf).is_ok() {
                    #[cfg(debug_assertions)]
                    eprintln!("[DEBUG] Probe: write-with-report-id succeeded");
                    self.protocol = Protocol::Legacy(WriteMethod::WriteWithReportId);
                    return Ok(());
                }

                #[cfg(debug_assertions)]
                eprintln!("[DEBUG] All write methods failed");
                Err(error::Error::HidError(hidapi::HidError::HidApiError {
                    message: "All write methods failed (feature_report, write, write+report_id)".into(),
                }))
            }
        }
    }

    pub fn set_effect(&mut self, effect: BaseEffects) -> Result<()> {
        self.current_state.effect_type = effect;
        self.refresh()?;

        Ok(())
    }

    pub fn set_speed(&mut self, speed: u8) -> Result<()> {
        if !SPEED_RANGE.contains(&speed) {
            return Err(RangeError { kind: RangeErrorKind::Speed }.into());
        }

        self.current_state.speed = speed;
        self.refresh()?;

        Ok(())
    }

    pub fn set_brightness(&mut self, brightness: u8) -> Result<()> {
        if self.is_dynamic_lighting() {
            // Map legacy 1-2 scale to percentage for WDL devices
            self.current_state.brightness = match brightness {
                1 => 50,
                2 => 100,
                _ => brightness.clamp(1, 100),
            };
        } else {
            if !BRIGHTNESS_RANGE.contains(&brightness) {
                return Err(RangeError { kind: RangeErrorKind::Brightness }.into());
            }
            self.current_state.brightness = brightness;
        }
        self.refresh()?;

        Ok(())
    }

    /// Set brightness as a percentage (1-100). Works for all device types.
    /// Legacy devices: 1-50 maps to Low, 51-100 maps to High.
    pub fn set_brightness_percent(&mut self, percent: u8) -> Result<()> {
        let percent = percent.clamp(1, 100);
        if self.is_dynamic_lighting() {
            self.current_state.brightness = percent;
        } else {
            self.current_state.brightness = if percent > 50 { 2 } else { 1 };
        }
        self.refresh()?;

        Ok(())
    }

    pub fn set_zone_by_index(&mut self, zone_index: u8, new_values: [u8; 3]) -> Result<()> {
        if !ZONE_RANGE.contains(&zone_index) {
            return Err(RangeError { kind: RangeErrorKind::Zone }.into());
        }

        for (i, _) in new_values.iter().enumerate() {
            let full_index = (zone_index * 3 + i as u8) as usize;
            self.current_state.rgb_values[full_index] = new_values[i];
        }
        self.refresh()?;

        Ok(())
    }

    pub fn set_colors_to(&mut self, new_values: &[u8; 12]) -> Result<()> {
        if self.is_dynamic_lighting() || matches!(self.current_state.effect_type, BaseEffects::Static | BaseEffects::Breath) {
            for (i, _) in new_values.iter().enumerate() {
                self.current_state.rgb_values[i] = new_values[i];
            }
            self.refresh()?;
        }

        Ok(())
    }

    pub fn solid_set_colors_to(&mut self, new_values: [u8; 3]) -> Result<()> {
        if self.is_dynamic_lighting() || matches!(self.current_state.effect_type, BaseEffects::Static | BaseEffects::Breath) {
            for i in (0..12).step_by(3) {
                self.current_state.rgb_values[i] = new_values[0];
                self.current_state.rgb_values[i + 1] = new_values[1];
                self.current_state.rgb_values[i + 2] = new_values[2];
            }
            self.refresh()?;
        }

        Ok(())
    }

    pub fn transition_colors_to(&mut self, target_colors: &[u8; 12], steps: u8, delay_between_steps: u64) -> Result<()> {
        if self.is_dynamic_lighting() || matches!(self.current_state.effect_type, BaseEffects::Static | BaseEffects::Breath) {
            let mut new_values = self.current_state.rgb_values.map(f32::from);
            let mut color_differences: [f32; 12] = [0.0; 12];
            for index in 0..12 {
                color_differences[index] = (f32::from(target_colors[index]) - f32::from(self.current_state.rgb_values[index])) / f32::from(steps);
            }
            if !self.stop_signal.load(Ordering::SeqCst) {
                for _step_num in 1..=steps {
                    if self.stop_signal.load(Ordering::SeqCst) {
                        break;
                    }
                    for (index, _) in color_differences.iter().enumerate() {
                        new_values[index] += color_differences[index];
                    }
                    self.current_state.rgb_values = new_values.map(|val| val as u8);

                    self.refresh()?;
                    thread::sleep(Duration::from_millis(delay_between_steps));
                }
                self.set_colors_to(target_colors)?;
            }
        }

        Ok(())
    }

    /// Returns true if this keyboard uses Windows Dynamic Lighting (WinRT API).
    pub fn is_dynamic_lighting(&self) -> bool {
        #[cfg(target_os = "windows")]
        {
            matches!(
                self.protocol,
                Protocol::WindowsDynamicLighting { .. } | Protocol::LampArrayHidManaged
            )
        }
        #[cfg(not(target_os = "windows"))]
        { false }
    }
}

/// Returns true if the given (VID, PID) pair is in our known device table.
fn is_known_pid(vid: u16, pid: u16) -> bool {
    KNOWN_DEVICE_INFOS.iter().any(|k| k.0 == vid && k.1 == pid)
}

/// Parse a HID report descriptor to extract LampArray report IDs.
///
/// Scans for items in the Lighting and Illumination usage page (0x59)
/// and maps known usage IDs to their associated HID report IDs.
fn parse_lamp_array_report_ids(descriptor: &[u8]) -> Option<LampArrayReportIds> {
    let mut ids = LampArrayReportIds {
        attributes: 0,
        lamp_attr_request: 0,
        lamp_attr_response: 0,
        multi_update: 0,
        range_update: 0,
        control: 0,
    };

    let mut usage_page: u32 = 0;
    let mut report_id: u8 = 0;
    let mut pending_usage: Option<u32> = None;
    let mut found_any = false;

    let mut i = 0;
    while i < descriptor.len() {
        let prefix = descriptor[i];

        // Long item (prefix 0xFE) — skip
        if prefix == 0xFE {
            if i + 2 >= descriptor.len() {
                break;
            }
            let data_size = descriptor[i + 1] as usize;
            i += 3 + data_size;
            continue;
        }

        let size = match prefix & 0x03 {
            0 => 0usize,
            1 => 1,
            2 => 2,
            3 => 4, // size encoding 3 means 4 data bytes
            _ => unreachable!(),
        };
        let item_type = (prefix >> 2) & 0x03;
        let tag = (prefix >> 4) & 0x0F;

        if i + 1 + size > descriptor.len() {
            break;
        }

        let value = match size {
            1 => descriptor[i + 1] as u32,
            2 => u16::from_le_bytes([descriptor[i + 1], descriptor[i + 2]]) as u32,
            4 => u32::from_le_bytes([
                descriptor[i + 1],
                descriptor[i + 2],
                descriptor[i + 3],
                descriptor[i + 4],
            ]),
            _ => 0,
        };

        match item_type {
            1 => {
                // Global item
                match tag {
                    0 => usage_page = value, // Usage Page
                    8 => {
                        report_id = value as u8; // Report ID
                        // Usage often appears before Report ID in this descriptor.
                        if let Some(usage) = pending_usage.take() {
                            if assign_lamp_array_report_id(&mut ids, usage, report_id) {
                                found_any = true;
                            }
                        }
                    }
                    _ => {}
                }
            }
            2 => {
                // Local item — Usage (tag 0)
                if tag == 0 && usage_page == LAMP_ARRAY_USAGE_PAGE as u32 {
                    pending_usage = Some(value);
                }
            }
            _ => {}
        }

        i += 1 + size;
    }

    if found_any {
        Some(normalize_lamp_array_report_ids(ids))
    } else {
        None
    }
}

fn assign_lamp_array_report_id(ids: &mut LampArrayReportIds, usage: u32, report_id: u8) -> bool {
    match usage {
        v if v == USAGE_LAMP_ARRAY_ATTRIBUTES_REPORT => ids.attributes = report_id,
        v if v == USAGE_LAMP_ATTRIBUTES_REQUEST_REPORT => ids.lamp_attr_request = report_id,
        v if v == USAGE_LAMP_ATTRIBUTES_RESPONSE_REPORT => ids.lamp_attr_response = report_id,
        v if v == USAGE_LAMP_MULTI_UPDATE_REPORT => ids.multi_update = report_id,
        v if v == USAGE_LAMP_RANGE_UPDATE_REPORT => ids.range_update = report_id,
        v if v == USAGE_LAMP_ARRAY_CONTROL_REPORT => ids.control = report_id,
        _ => return false,
    }
    true
}

/// ITE 8258 / HID LampArray spec uses reports 4/5/6. An older parser bound
/// usages to the previous report ID and then used Output writes, so click-away
/// HID never actually painted this keyboard.
fn normalize_lamp_array_report_ids(mut ids: LampArrayReportIds) -> LampArrayReportIds {
    if ids.multi_update == 3 && ids.range_update == 4 && ids.control == 5 {
        ids.multi_update = 4;
        ids.range_update = 5;
        ids.control = 6;
    }
    if ids.multi_update == 0 {
        ids.multi_update = 4;
    }
    if ids.range_update == 0 {
        ids.range_update = 5;
    }
    if ids.control == 0 {
        ids.control = 6;
    }
    ids
}

/// Try to open a HID LampArray interface for a known device.
/// Returns a Keyboard configured for the LampArray protocol, or None.
#[cfg(target_os = "windows")]
fn try_lamp_array_keyboard(api: &HidApi, stop_signal: &Arc<AtomicBool>) -> Option<Keyboard> {
    log_hid_inventory(api, "try_lamp_array_keyboard");
    log_to_file("try_lamp_array_keyboard: searching for usage_page=0x0059 usage=0x0001...");
    let lamp_info = api.device_list().find(|d| {
        is_known_pid(d.vendor_id(), d.product_id())
            && d.usage_page() == LAMP_ARRAY_USAGE_PAGE
            && d.usage() == LAMP_ARRAY_USAGE
    });
    let lamp_info = match lamp_info {
        Some(i) => i,
        None => {
            log_to_file("try_lamp_array_keyboard: NO device found with LampArray usage page");
            return None;
        }
    };

    log_to_file(&format!(
        "try_lamp_array_keyboard: found VID={:#06x} PID={:#06x} usage_page={:#06x} — opening...",
        lamp_info.vendor_id(), lamp_info.product_id(), lamp_info.usage_page()
    ));
    if reverse_engineering_mode_enabled() {
        log_to_file(&format!(
            "RE-MODE [try_lamp_array_keyboard]: selected iface={} usage={:#06x} path={}",
            lamp_info.interface_number(),
            lamp_info.usage(),
            lamp_info.path().to_string_lossy()
        ));
    }

    let device = match lamp_info.open_device(api) {
        Ok(d) => d,
        Err(e) => {
            log_to_file(&format!("try_lamp_array_keyboard: failed to open device: {}", e));
            return None;
        }
    };

    // Parse the report descriptor to discover LampArray report IDs
    let mut desc_buf = [0u8; 4096];
    let desc_len = match device.get_report_descriptor(&mut desc_buf) {
        Ok(n) => n,
        Err(e) => {
            log_to_file(&format!("try_lamp_array_keyboard: failed to get report descriptor: {}", e));
            return None;
        }
    };
    log_descriptor_preview("try_lamp_array_keyboard", &desc_buf[..desc_len]);
    let report_ids = match parse_lamp_array_report_ids(&desc_buf[..desc_len]) {
        Some(ids) => ids,
        None => {
            log_to_file("try_lamp_array_keyboard: failed to parse LampArray report IDs from descriptor");
            return None;
        }
    };

    log_to_file(&format!(
        "try_lamp_array_keyboard: report IDs — attrs:{}, multi:{}, range:{}, ctrl:{}",
        report_ids.attributes, report_ids.multi_update, report_ids.range_update, report_ids.control
    ));

    let is_loq_15irx10 = lamp_info.product_id() == 0xc693;
    let lamp_count = if is_loq_15irx10 {
        log_to_file("try_lamp_array_keyboard: LOQ 15IRX10 detected, using known lamp_count=24 and raw HID managed mode");
        24
    } else {
        let mut attr_buf = [0u8; 64];
        attr_buf[0] = report_ids.attributes;
        let n = match device.get_feature_report(&mut attr_buf) {
            Ok(n) => n,
            Err(e) => {
                log_to_file(&format!("try_lamp_array_keyboard: failed to get LampArray attributes: {}", e));
                return None;
            }
        };
        if n < 3 {
            log_to_file(&format!("try_lamp_array_keyboard: attributes report too short: {} bytes", n));
            return None;
        }
        let lamp_count = u16::from_le_bytes([attr_buf[1], attr_buf[2]]);
        if lamp_count == 0 {
            log_to_file("try_lamp_array_keyboard: lamp count is 0");
            return None;
        }
        lamp_count
    };

    log_to_file(&format!("try_lamp_array_keyboard: {} lamps detected", lamp_count));

    // Take host control (autonomous_mode = 0)
    let control_buf = [report_ids.control, 0x00];
    if let Err(e) = device.send_feature_report(&control_buf) {
        log_to_file(&format!("try_lamp_array_keyboard: failed to take host control: {}", e));
        return None;
    }
    log_to_file("try_lamp_array_keyboard: host control acquired");

    let current_state = LightingState {
        effect_type: BaseEffects::Static,
        speed: 1,
        brightness: if is_loq_15irx10 { 50 } else { 1 },
        rgb_values: [0; 12],
    };

    #[cfg(target_os = "windows")]
    let raw_keepalive_state = if is_loq_15irx10 {
        Some(Arc::new(std::sync::Mutex::new(WinLampColors {
            rgb_values: [0; 12],
            brightness: 50,
            speed: 1,
            effect_type: BaseEffects::Static,
            lamp_count,
        })))
    } else {
        None
    };

    let window_active = Arc::new(AtomicBool::new(true));

    #[cfg(target_os = "windows")]
    let (raw_keepalive, keyboard_device) = if is_loq_15irx10 {
        let keepalive = spawn_raw_hid_keepalive(
            stop_signal.clone(),
            raw_keepalive_state.clone().unwrap(),
            lamp_count,
            device,
            report_ids,
            window_active.clone(),
        );
        (Some(keepalive), None)
    } else {
        (None, Some(device))
    };

    Some(Keyboard {
        keyboard_hid: keyboard_device,
        #[cfg(target_os = "windows")]
        win_lamp_array: Arc::new(std::sync::Mutex::new(None)),
        #[cfg(target_os = "windows")]
        win_color_state: raw_keepalive_state,
        #[cfg(target_os = "windows")]
        _win_playlist: None,
        #[cfg(target_os = "windows")]
        _win_keepalive: raw_keepalive,
        current_state,
        stop_signal: stop_signal.clone(),
        window_active,
        protocol: if is_loq_15irx10 {
            Protocol::LampArrayHidManaged
        } else {
            Protocol::LampArrayHid { report_ids, lamp_count }
        },
    })
}

/// LOQ-specific fallback: open the vendor endpoint directly and use legacy payload writes.
/// This path is useful when LampArray path is present but does not visibly update LEDs.
#[cfg(target_os = "windows")]
fn try_loq_vendor_keyboard(api: &HidApi, stop_signal: &Arc<AtomicBool>) -> Option<Keyboard> {
    log_to_file("try_loq_vendor_keyboard: searching usage_page=0xff89 usage=0x00cc...");
    let info = api.device_list().find(|d| {
        d.vendor_id() == 0x048d
            && d.product_id() == 0xc693
            && d.usage_page() == 0xff89
            && d.usage() == 0x00cc
    })?;

    log_to_file(&format!(
        "try_loq_vendor_keyboard: found iface={} path={} — opening...",
        info.interface_number(),
        info.path().to_string_lossy()
    ));

    let device = match info.open_device(api) {
        Ok(d) => d,
        Err(e) => {
            log_to_file(&format!("try_loq_vendor_keyboard: failed to open: {}", e));
            return None;
        }
    };

    let current_state = LightingState {
        effect_type: BaseEffects::Static,
        speed: 1,
        brightness: 1,
        rgb_values: [0; 12],
    };

    let mut keyboard = Keyboard {
        keyboard_hid: Some(device),
        #[cfg(target_os = "windows")]
        win_lamp_array: Arc::new(std::sync::Mutex::new(None)),
        #[cfg(target_os = "windows")]
        win_color_state: None,
        #[cfg(target_os = "windows")]
        _win_playlist: None,
        #[cfg(target_os = "windows")]
        _win_keepalive: None,
        current_state,
        stop_signal: stop_signal.clone(),
        window_active: Arc::new(AtomicBool::new(true)),
        protocol: Protocol::Legacy(WriteMethod::Unknown),
    };

    if let Err(e) = keyboard.refresh() {
        log_to_file(&format!("try_loq_vendor_keyboard: initial refresh failed: {}", e));
        return None;
    }

    log_to_file("try_loq_vendor_keyboard: SUCCESS");
    Some(keyboard)
}

/// Try to open the raw HID LampArray interface for direct writes.
/// Called from the keepalive thread to get a secondary HID handle that
/// bypasses the WDL arbitrator's focus-based priority.
/// `known_lamp_count` is provided from the WDL path so we skip the
/// attributes feature report (report ID 0 is broken on Windows HID API).
#[cfg(target_os = "windows")]
fn open_raw_lamp_array_hid(known_lamp_count: u16) -> Option<(HidDevice, LampArrayReportIds, u16)> {
    let api = HidApi::new().ok()?;
    log_hid_inventory(&api, "open_raw_lamp_array_hid");
    let lamp_info = api.device_list().find(|d| {
        is_known_pid(d.vendor_id(), d.product_id())
            && d.usage_page() == LAMP_ARRAY_USAGE_PAGE
            && d.usage() == LAMP_ARRAY_USAGE
    })?;
    if reverse_engineering_mode_enabled() {
        log_to_file(&format!(
            "RE-MODE [open_raw_lamp_array_hid]: selected iface={} path={}",
            lamp_info.interface_number(),
            lamp_info.path().to_string_lossy()
        ));
    }
    let device = match lamp_info.open_device(&api) {
        Ok(d) => d,
        Err(e) => {
            log_to_file(&format!("open_raw_lamp_array_hid: failed to open device: {}", e));
            return None;
        }
    };

    let mut desc_buf = [0u8; 4096];
    let desc_len = match device.get_report_descriptor(&mut desc_buf) {
        Ok(n) => n,
        Err(e) => {
            log_to_file(&format!("open_raw_lamp_array_hid: failed to get descriptor: {}", e));
            return None;
        }
    };
    log_descriptor_preview("open_raw_lamp_array_hid", &desc_buf[..desc_len]);
    let report_ids = match parse_lamp_array_report_ids(&desc_buf[..desc_len]) {
        Some(ids) => ids,
        None => {
            log_to_file("open_raw_lamp_array_hid: failed to parse report IDs");
            return None;
        }
    };

    log_to_file(&format!(
        "open_raw_lamp_array_hid: report IDs — multi:{}, range:{}, ctrl:{} (using known lamp_count={})",
        report_ids.multi_update, report_ids.range_update, report_ids.control, known_lamp_count
    ));

    // Take host control (autonomous_mode = 0) via LampArray Control report.
    if report_ids.control != 0 {
        match send_lamp_array_feature(&device, &[report_ids.control, 0x00]) {
            Ok(_) => log_to_file(&format!(
                "open_raw_lamp_array_hid: host control acquired ctrl={} multi={} range={}",
                report_ids.control, report_ids.multi_update, report_ids.range_update
            )),
            Err(e) => log_to_file(&format!("open_raw_lamp_array_hid: host control failed: {}", e)),
        }
    }

    Some((device, report_ids, known_lamp_count))
}

#[cfg(target_os = "windows")]
fn send_lamp_array_feature(device: &HidDevice, payload: &[u8]) -> std::result::Result<(), hidapi::HidError> {
    let mut buf = [0u8; 65];
    let n = payload.len().min(buf.len());
    buf[..n].copy_from_slice(&payload[..n]);
    device.send_feature_report(&buf).map(|_| ())
}

/// Paint every lamp by expanding the 4 zone colors. This keyboard only accepts
/// HID LampArray *feature* reports (IDs 4/5/6). Output writes are ignored.
#[cfg(target_os = "windows")]
fn write_lamp_array_zones(
    device: &HidDevice,
    report_ids: &LampArrayReportIds,
    lamp_count: u16,
    rgb: &[u8; 12],
    intensity: u8,
) {
    let lc = lamp_count.max(1) as usize;
    let zones = std::cmp::min(4, lc);
    let intensity = if intensity == 0 { 1 } else { intensity };

    if report_ids.range_update != 0 {
        for z in 0..zones {
            let start = (z * lc) / zones;
            let end_excl = ((z + 1) * lc) / zones;
            if start >= end_excl {
                continue;
            }
            let start_id = start as u16;
            let end_id = (end_excl - 1) as u16;
            let buf = [
                report_ids.range_update,
                if z + 1 == zones { 0x01 } else { 0x00 },
                start_id as u8,
                (start_id >> 8) as u8,
                end_id as u8,
                (end_id >> 8) as u8,
                rgb[z * 3],
                rgb[z * 3 + 1],
                rgb[z * 3 + 2],
                intensity,
            ];
            let _ = send_lamp_array_feature(device, &buf);
        }
        return;
    }

    let mut lamp = 0usize;
    while lamp < lc {
        let chunk = std::cmp::min(8, lc - lamp);
        let complete = lamp + chunk >= lc;
        let mut buf = [0u8; 51];
        buf[0] = report_ids.multi_update;
        buf[1] = chunk as u8;
        buf[2] = if complete { 0x01 } else { 0x00 };
        for i in 0..chunk {
            let id = (lamp + i) as u16;
            buf[3 + i * 2] = id as u8;
            buf[4 + i * 2] = (id >> 8) as u8;
            let z = std::cmp::min((lamp + i) * zones / lc, zones.saturating_sub(1));
            let co = 19 + i * 4;
            buf[co] = rgb[z * 3];
            buf[co + 1] = rgb[z * 3 + 1];
            buf[co + 2] = rgb[z * 3 + 2];
            buf[co + 3] = intensity;
        }
        let _ = send_lamp_array_feature(device, &buf);
        lamp += chunk;
    }
}

#[cfg(target_os = "windows")]
#[allow(dead_code)]
fn write_vendor_unfocused_rgb(targets: &mut [VendorHidTarget], rgb: &[u8; 12], brightness: u8) {
    let legacy_brightness: u8 = if brightness > 50 { 2 } else { 1 };
    let mut payload = [0u8; 33];
    payload[0] = 0xcc;
    payload[1] = 0x16;
    payload[2] = 0x01;
    payload[3] = 1;
    payload[4] = legacy_brightness;
    payload[5..17].copy_from_slice(rgb);
    for target in targets {
        let is_led_endpoint = target.usage_page == 0xff89 && target.usage == 0x00cc;
        if !is_led_endpoint {
            continue;
        }
        let _ = target.device.send_feature_report(&payload);
        let _ = target.device.write(&payload);
    }
}

/// Try to open all vendor-like HID interfaces on MI_00.
/// Some LOQ firmwares expose multiple vendor collections, and only one of them
/// may actually react to RGB payloads.
#[cfg(target_os = "windows")]
fn open_vendor_hids() -> Vec<VendorHidTarget> {
    let api = match HidApi::new() {
        Ok(api) => api,
        Err(_) => return Vec::new(),
    };
    log_hid_inventory(&api, "open_vendor_hid");

    let mut devices = Vec::new();
    for info in api.device_list().filter(|d| {
        is_known_pid(d.vendor_id(), d.product_id())
            && d.interface_number() == 0
            && d.usage_page() >= 0xff00
    }) {
        let label = format!(
            "usage_page={:#06x} usage={:#06x} iface={} path={}",
            info.usage_page(),
            info.usage(),
            info.interface_number(),
            info.path().to_string_lossy()
        );
        if reverse_engineering_mode_enabled() {
            log_to_file(&format!("RE-MODE [open_vendor_hid]: candidate {}", label));
        }

        let path = info.path().to_owned();
        match api.open_path(&path) {
            Ok(device) => {
                let mut descriptor = [0u8; 4096];
                let report_ids = match device.get_report_descriptor(&mut descriptor) {
                    Ok(len) => {
                        let context = format!("vendor {}", label);
                        log_descriptor_preview(&context, &descriptor[..len]);
                        log_report_id_sets(&context, &descriptor[..len])
                    }
                    Err(e) => {
                        log_to_file(&format!(
                            "open_vendor_hid: failed to read report descriptor for {}: {}",
                            label, e
                        ));
                        HidReportIdSets::default()
                    }
                };
                let context = format!("vendor {}", label);
                for (report_id, report_len) in &report_ids.feature_lengths {
                    log_feature_report_snapshot(&device, &context, *report_id, *report_len);
                }
                let gen7_controller = if info.usage_page() == 0xff89
                    && report_ids.feature.contains(&GEN7_REPORT_ID)
                    && (info.usage() == 0x0007 || info.product_id() == 0xc693)
                {
                    try_load_gen7_controller(&device, &label, info.product_id())
                        .or_else(|| fallback_loq_gen7_controller(info.product_id()))
                } else {
                    None
                };
                devices.push(VendorHidTarget {
                    label,
                    product_id: info.product_id(),
                    usage_page: info.usage_page(),
                    usage: info.usage(),
                    device,
                    output_report_ids: report_ids.output,
                    feature_report_ids: report_ids.feature,
                    feature_report_lengths: report_ids.feature_lengths,
                    gen7_controller,
                    mode: VendorWriteMode::Probe,
                });
            }
            Err(e) => log_to_file(&format!("open_vendor_hid: failed to open {}: {}", label, e)),
        }
    }

    devices
}

#[cfg(target_os = "windows")]
fn spawn_raw_hid_keepalive(
    _stop_signal: Arc<AtomicBool>,
    shared_state: Arc<std::sync::Mutex<WinLampColors>>,
    lamp_count: u16,
    device: HidDevice,
    report_ids: LampArrayReportIds,
    window_active: Arc<AtomicBool>,
) -> std::thread::JoinHandle<()> {
    thread::spawn(move || {
        let mut tick: u64 = 0;
        let mut vendor_hids = open_vendor_hids();
        let vendor_primary = loq_vendor_primary_enabled();
        let force_vendor_only = force_vendor_only_enabled();

        log_to_file(&format!(
            "Raw HID keepalive: using owned device with multi_update_id={}, lamps={}",
            report_ids.multi_update, lamp_count
        ));
        let skip_lamparray = force_vendor_only;
        if force_vendor_only {
            log_to_file("Raw HID keepalive: LEGION_RGB_FORCE_VENDOR_ONLY enabled (skipping MI_01 LampArray writes)");
        }
        if vendor_primary {
            log_to_file("Raw HID keepalive: LOQ vendor-primary mode enabled");
        }
        if vendor_hids.is_empty() {
            log_to_file("Raw HID keepalive: vendor HID interface NOT available");
        } else {
            log_to_file(&format!(
                "Raw HID keepalive: vendor HID interfaces OPENED ({})",
                vendor_hids.len()
            ));
            for target in &vendor_hids {
                log_to_file(&format!(
                    "Raw HID keepalive: vendor candidate {} output_ids={:?} feature_ids={:?} feature_lengths={:?}",
                    target.label, target.output_report_ids, target.feature_report_ids, target.feature_report_lengths
                ));
            }
        }

        let mut last_non_black_rgb = [0u8; 12];
        loop {
            let (rgb_values, brightness, speed, effect_type, lamp_count) = {
                let state = shared_state.lock().unwrap();
                (
                    state.rgb_values,
                    state.brightness,
                    state.speed,
                    state.effect_type,
                    state.lamp_count,
                )
            };
            if is_vivid_rgb_frame(&rgb_values) {
                last_non_black_rgb = rgb_values;
            }
            let rgb_values = if is_vivid_rgb_frame(&rgb_values) {
                rgb_values
            } else {
                last_non_black_rgb
            };

            let window_is_active = window_active.load(Ordering::Relaxed);
            let _ = window_is_active;

            if !skip_lamparray {
                let intensity = ((brightness.clamp(1, 100) as u16 * 255) / 100) as u8;
                write_lamp_array_zones(&device, &report_ids, lamp_count, &rgb_values, intensity);

                if report_ids.control != 0 {
                    let ctrl = [report_ids.control, 0x00];
                    let _ = device.send_feature_report(&ctrl);
                    let _ = device.write(&ctrl);
                }
            }

            if !vendor_hids.is_empty() {
                let legacy_brightness: u8 = if brightness > 50 { 2 } else { 1 };
                let mut payload = [0u8; 33];
                payload[0] = 0xcc;
                payload[1] = 0x16;
                payload[2] = 0x01;
                payload[3] = 1;
                payload[4] = legacy_brightness;
                payload[5..17].copy_from_slice(&rgb_values);

                for target in &mut vendor_hids {
                    if let Some(mut controller) = target.gen7_controller.take() {
                        apply_gen7_controller(
                            target,
                            &mut controller,
                            rgb_values,
                            brightness,
                            speed,
                            effect_type,
                            tick,
                        );
                        target.gen7_controller = Some(controller);
                        continue;
                    }

                    let is_led_endpoint = target.usage_page == 0xff89 && target.usage == 0x00cc;
                    match target.mode.clone() {
                        VendorWriteMode::Disabled => {}
                        VendorWriteMode::FeatureReport => {
                            if !is_led_endpoint {
                                continue;
                            }
                            if let Err(e) = target.device.send_feature_report(&payload) {
                                if reverse_engineering_mode_enabled() {
                                    log_to_file(&format!(
                                        "RE-MODE [raw_keepalive]: vendor feature report failed on {}: {}",
                                        target.label, e
                                    ));
                                }
                                target.mode = VendorWriteMode::Probe;
                            }
                        }
                        VendorWriteMode::FeatureWithZeroReportId => {
                            if !is_led_endpoint {
                                continue;
                            }
                            let buf = prefixed_payload(0x00, &payload);
                            if let Err(e) = target.device.send_feature_report(&buf) {
                                if reverse_engineering_mode_enabled() {
                                    log_to_file(&format!(
                                        "RE-MODE [raw_keepalive]: vendor feature+0x00 failed on {}: {}",
                                        target.label, e
                                    ));
                                }
                                target.mode = VendorWriteMode::Probe;
                            }
                        }
                        VendorWriteMode::FeatureWithReportId(report_id) => {
                            if !is_led_endpoint {
                                continue;
                            }
                            let buf = prefixed_payload(report_id, &payload);
                            if let Err(e) = target.device.send_feature_report(&buf) {
                                if reverse_engineering_mode_enabled() {
                                    log_to_file(&format!(
                                        "RE-MODE [raw_keepalive]: vendor feature+id {:#04x} failed on {}: {}",
                                        report_id, target.label, e
                                    ));
                                }
                                target.mode = VendorWriteMode::Probe;
                            }
                        }
                        VendorWriteMode::Write => {
                            if !is_led_endpoint {
                                continue;
                            }
                            if let Err(e) = target.device.write(&payload) {
                                if reverse_engineering_mode_enabled() {
                                    log_to_file(&format!(
                                        "RE-MODE [raw_keepalive]: vendor write failed on {}: {}",
                                        target.label, e
                                    ));
                                }
                                target.mode = VendorWriteMode::Probe;
                            }
                        }
                        VendorWriteMode::WriteWithZeroReportId => {
                            if !is_led_endpoint {
                                continue;
                            }
                            let buf = prefixed_payload(0x00, &payload);
                            if let Err(e) = target.device.write(&buf) {
                                if reverse_engineering_mode_enabled() {
                                    log_to_file(&format!(
                                        "RE-MODE [raw_keepalive]: vendor write+0x00 failed on {}: {}",
                                        target.label, e
                                    ));
                                }
                                target.mode = VendorWriteMode::Probe;
                            }
                        }
                        VendorWriteMode::WriteWithReportId(report_id) => {
                            if !is_led_endpoint {
                                continue;
                            }
                            let buf = prefixed_payload(report_id, &payload);
                            if let Err(e) = target.device.write(&buf) {
                                if reverse_engineering_mode_enabled() {
                                    log_to_file(&format!(
                                        "RE-MODE [raw_keepalive]: vendor write+id {:#04x} failed on {}: {}",
                                        report_id, target.label, e
                                    ));
                                }
                                target.mode = VendorWriteMode::Probe;
                            }
                        }
                        VendorWriteMode::Probe => {
                            // Probe each vendor collection independently. On LOQ firmwares,
                            // one collection may ACK while another is the one that really drives LEDs.
                            if target.device.write(&payload).is_ok() {
                                if is_led_endpoint {
                                    target.mode = VendorWriteMode::Write;
                                    log_to_file(&format!(
                                        "Raw HID keepalive: vendor HID write probe on {}: Write works",
                                        target.label
                                    ));
                                } else {
                                    target.mode = VendorWriteMode::Disabled;
                                    log_to_file(&format!(
                                        "Raw HID keepalive: vendor HID write probe on {}: Write works but endpoint is not treated as LED stream",
                                        target.label
                                    ));
                                }
                                continue;
                            }

                            let buf = prefixed_payload(0x00, &payload);
                            if target.device.write(&buf).is_ok() {
                                if is_led_endpoint {
                                    target.mode = VendorWriteMode::WriteWithZeroReportId;
                                    log_to_file(&format!(
                                        "Raw HID keepalive: vendor HID write probe on {}: WriteWithZeroReportId works",
                                        target.label
                                    ));
                                } else {
                                    target.mode = VendorWriteMode::Disabled;
                                    log_to_file(&format!(
                                        "Raw HID keepalive: vendor HID write probe on {}: WriteWithZeroReportId works but endpoint is not treated as LED stream",
                                        target.label
                                    ));
                                }
                                continue;
                            }

                            if target.device.send_feature_report(&payload).is_ok() {
                                if is_led_endpoint {
                                    target.mode = VendorWriteMode::FeatureReport;
                                    log_to_file(&format!(
                                        "Raw HID keepalive: vendor HID write probe on {}: FeatureReport works",
                                        target.label
                                    ));
                                } else {
                                    target.mode = VendorWriteMode::Disabled;
                                    log_to_file(&format!(
                                        "Raw HID keepalive: vendor HID write probe on {}: FeatureReport works but endpoint is not treated as LED stream",
                                        target.label
                                    ));
                                }
                                continue;
                            }

                            let buf = prefixed_payload(0x00, &payload);
                            if target.device.send_feature_report(&buf).is_ok() {
                                if is_led_endpoint {
                                    target.mode = VendorWriteMode::FeatureWithZeroReportId;
                                    log_to_file(&format!(
                                        "Raw HID keepalive: vendor HID write probe on {}: FeatureReportWithZeroReportId works",
                                        target.label
                                    ));
                                } else {
                                    target.mode = VendorWriteMode::Disabled;
                                    log_to_file(&format!(
                                        "Raw HID keepalive: vendor HID write probe on {}: FeatureReportWithZeroReportId works but endpoint is not treated as LED stream",
                                        target.label
                                    ));
                                }
                                continue;
                            }

                            let mut matched = false;
                            for report_id in target.output_report_ids.iter().copied().filter(|id| *id != 0) {
                                let buf = prefixed_payload(report_id, &payload);
                                if target.device.write(&buf).is_ok() {
                                    if is_led_endpoint {
                                        target.mode = VendorWriteMode::WriteWithReportId(report_id);
                                        log_to_file(&format!(
                                            "Raw HID keepalive: vendor HID write probe on {}: WriteWithReportId({:#04x}) works",
                                            target.label, report_id
                                        ));
                                    } else {
                                        target.mode = VendorWriteMode::Disabled;
                                        log_to_file(&format!(
                                            "Raw HID keepalive: vendor HID write probe on {}: WriteWithReportId({:#04x}) works but endpoint is not treated as LED stream",
                                            target.label, report_id
                                        ));
                                    }
                                    matched = true;
                                    break;
                                }
                            }
                            if matched {
                                continue;
                            }

                            for report_id in target.feature_report_ids.iter().copied().filter(|id| *id != 0) {
                                let buf = prefixed_payload(report_id, &payload);
                                if target.device.send_feature_report(&buf).is_ok() {
                                    if is_led_endpoint {
                                        target.mode = VendorWriteMode::FeatureWithReportId(report_id);
                                        log_to_file(&format!(
                                            "Raw HID keepalive: vendor HID write probe on {}: FeatureReportWithReportId({:#04x}) works",
                                            target.label, report_id
                                        ));
                                    } else {
                                        target.mode = VendorWriteMode::Disabled;
                                        log_to_file(&format!(
                                            "Raw HID keepalive: vendor HID write probe on {}: FeatureReportWithReportId({:#04x}) works but endpoint is not treated as LED stream",
                                            target.label, report_id
                                        ));
                                    }
                                    matched = true;
                                    break;
                                }
                            }
                            if matched {
                                continue;
                            }

                            target.mode = VendorWriteMode::Disabled;
                            log_to_file(&format!(
                                "Raw HID keepalive: vendor HID write probe on {}: ALL methods failed; disabling endpoint",
                                target.label
                            ));
                        }
                    }
                }
            }

            tick += 1;
            if tick % 800 == 0 {
                log_to_file(&format!(
                    "Raw HID keepalive: tick={}, lamp_count={}, brightness={}, vendor_hid={}, rgb=[{},{},{},{},{},{},{},{},{},{},{},{}]",
                    tick,
                    lamp_count,
                    brightness,
                    !vendor_hids.is_empty(),
                    rgb_values[0], rgb_values[1], rgb_values[2],
                    rgb_values[3], rgb_values[4], rgb_values[5],
                    rgb_values[6], rgb_values[7], rgb_values[8],
                    rgb_values[9], rgb_values[10], rgb_values[11]
                ));
            }

            thread::sleep(Duration::from_millis(30));
        }
    })
}

#[cfg(target_os = "windows")]
fn open_wdl_lamp_array(device_id: &str) -> Option<LampArray> {
    let id = HSTRING::from(device_id);
    match LampArray::FromIdAsync(&id) {
        Ok(op) => match op.get() {
            Ok(lamp_array) => Some(lamp_array),
            Err(e) => {
                log_to_file(&format!("WDL: reopen LampArray failed: {}", e));
                None
            }
        },
        Err(e) => {
            log_to_file(&format!("WDL: FromIdAsync failed: {}", e));
            None
        }
    }
}

/// Try to use the Windows Dynamic Lighting WinRT API.
/// This is the most reliable method for devices like the LOQ 15IRX10 that
/// support Windows Dynamic Lighting, because Windows already has exclusive
/// access to the HID LampArray interface.
#[cfg(target_os = "windows")]
fn try_windows_dynamic_lighting(stop_signal: &Arc<AtomicBool>) -> Option<Keyboard> {
    let selector = LampArray::GetDeviceSelector().ok()?;
    let devices = DeviceInformation::FindAllAsyncAqsFilter(&selector)
        .ok()?
        .get()
        .ok()?;

    let count = devices.Size().ok()?;

    log_to_file(&format!("WDL: found {} LampArray device(s)", count));
    #[cfg(debug_assertions)]
    eprintln!("[DEBUG] Windows Dynamic Lighting: found {} LampArray device(s)", count);

    for i in 0..count {
        let dev = match devices.GetAt(i) {
            Ok(d) => d,
            Err(_) => continue,
        };
        let id = match dev.Id() {
            Ok(id) => id,
            Err(_) => continue,
        };
        let device_id = id.to_string();
        let id_str = device_id.to_lowercase();

        // Check if this device matches any of our known VID/PID pairs
        let is_known = KNOWN_DEVICE_INFOS.iter().any(|&(vid, pid, _, _)| {
            let vid_str = format!("vid_{:04x}", vid);
            let pid_str = format!("pid_{:04x}", pid);
            id_str.contains(&vid_str) && id_str.contains(&pid_str)
        });
        if !is_known {
            #[cfg(debug_assertions)]
            eprintln!("[DEBUG]   Skipping non-Lenovo LampArray: {}", id_str);
            continue;
        }

        #[cfg(debug_assertions)]
        eprintln!("[DEBUG]   Opening Lenovo LampArray: {}", id_str);

        let lamp_array = match LampArray::FromIdAsync(&id) {
            Ok(op) => match op.get() {
                Ok(la) => la,
                Err(e) => {
                    #[cfg(debug_assertions)]
                    eprintln!("[DEBUG]   Failed to open LampArray: {}", e);
                    continue;
                }
            },
            Err(e) => {
                #[cfg(debug_assertions)]
                eprintln!("[DEBUG]   Failed to start FromIdAsync: {}", e);
                continue;
            }
        };

        let lamp_count = match lamp_array.LampCount() {
            Ok(c) => c as u16,
            Err(_) => continue,
        };
        if lamp_count == 0 {
            continue;
        }

        log_to_file(&format!("WDL: opened LampArray with {} lamps, id={}", lamp_count, id_str));
        #[cfg(debug_assertions)]
        eprintln!(
            "[DEBUG] Windows Dynamic Lighting: opened LampArray with {} lamps",
            lamp_count
        );

        // Shared color state: the effect callback reads, refresh() writes
        let shared_state = Arc::new(std::sync::Mutex::new(WinLampColors {
            rgb_values: [0; 12],
            brightness: 50,
            speed: 1,
            effect_type: BaseEffects::Static,
            lamp_count,
        }));

        let use_wdl_playlist = std::env::var("LEGION_RGB_WDL_USE_PLAYLIST")
            .ok()
            .as_deref()
            == Some("1");
        let callback_heartbeat = Arc::new(AtomicU64::new(0));

        let playlist = if use_wdl_playlist {
            // Create a custom effect covering ALL lamp indices
            let all_indices: Vec<i32> = (0..lamp_count as i32).collect();
            let effect = match LampArrayCustomEffect::CreateInstance(&lamp_array, &all_indices) {
                Ok(e) => e,
                Err(e) => {
                    #[cfg(debug_assertions)]
                    eprintln!("[DEBUG]   Failed to create LampArrayCustomEffect: {}", e);
                    continue;
                }
            };

            // Effect runs for a long time, callback fires at ~20fps
            let _ = effect.SetDuration(TimeSpan { Duration: 36_000_000_000 }); // 1 hour
            let _ = effect.SetUpdateInterval(TimeSpan { Duration: 500_000 }); // 50ms

            // Register callback that maps 4 color zones across ALL lamps
            let shared_clone = shared_state.clone();
            let callback_heartbeat_clone = callback_heartbeat.clone();
            let _ = effect.UpdateRequested(&TypedEventHandler::new(
                move |_effect, args: &Option<LampArrayUpdateRequestedEventArgs>| {
                    callback_heartbeat_clone.fetch_add(1, Ordering::Relaxed);
                    if let Some(args) = args {
                        let state = shared_clone.lock().unwrap();
                        if !is_vivid_rgb_frame(&state.rgb_values) {
                            // Avoid replaying black/dim frames from transient minimized-state updates.
                            return Ok(());
                        }
                        let lc = state.lamp_count as usize;
                        let zones = std::cmp::min(4, lc);
                        if zones == 0 {
                            return Ok(());
                        }

                        let intensity_factor = (state.brightness as f64 / 100.0).clamp(0.01, 1.0);

                        // Divide ALL lamps evenly into 4 zones
                        for z in 0..zones {
                            let zone_start = (z * lc) / zones;
                            let zone_end = ((z + 1) * lc) / zones;
                            if zone_start >= zone_end {
                                continue;
                            }
                            let r = (state.rgb_values[z * 3] as f64 * intensity_factor) as u8;
                            let g = (state.rgb_values[z * 3 + 1] as f64 * intensity_factor) as u8;
                            let b = (state.rgb_values[z * 3 + 2] as f64 * intensity_factor) as u8;
                            let color = Color { A: 255, R: r, G: g, B: b };
                            let indices: Vec<i32> =
                                (zone_start..zone_end).map(|i| i as i32).collect();
                            let _ = args.SetSingleColorForIndices(color, &indices);
                        }
                    }
                    Ok(())
                },
            ));

            // Start a persistent effect playlist — optional in reverse-engineering mode.
            let playlist = match LampArrayEffectPlaylist::new() {
                Ok(p) => p,
                Err(_) => continue,
            };
            let _ = playlist.Append(&effect);
            let _ = playlist.SetRepetitionMode(LampArrayRepetitionMode::Forever);
            let _ = playlist.OverrideZIndex(i32::MAX);
            let _ = playlist.Start();
            log_to_file("WDL: playlist mode enabled (LEGION_RGB_WDL_USE_PLAYLIST=1)");
            Some(playlist)
        } else {
            log_to_file("WDL: direct-write mode enabled (playlist disabled by default)");
            None
        };

        // Keepalive thread: continuously re-assert colors using multiple paths
        // to survive Windows Dynamic Lighting focus-loss override.
        let is_loq_15irx10 = id_str.contains("vid_048d") && id_str.contains("pid_c693");
        let window_active = Arc::new(AtomicBool::new(true));
        let shared_lamp = Arc::new(std::sync::Mutex::new(Some(lamp_array)));
        set_controlled_by_foreground_app(true);
        let keepalive_lamp_array = shared_lamp.clone();
        let reopen_device_id = device_id.clone();
        let keepalive_playlist = playlist.clone();
        let keepalive_state = shared_state.clone();
        let keepalive_callback_heartbeat = callback_heartbeat.clone();
        let keepalive_window_active = window_active.clone();
        let keepalive = thread::spawn(move || {
            let mut tick: u64 = 0;
            let mut last_non_black_rgb = [0u8; 12];
            let mut set_color_failures: u64 = 0;
            let mut last_callback_heartbeat: u64 = 0;
            let mut callback_stall_ticks: u64 = 0;
            let mut focused_ticks: u64 = 0;
            let mut unfocused_ticks: u64 = 0;
            let mut holding_lamp = true;
            let mut last_logged_focus: Option<bool> = None;
            let mut pending_ambient_off = false;
            let mut _ambient_hold_on = false;
            let mut last_hold_rgb = [0u8; 12];
            let mut rgb_unchanged_ticks: u64 = 0;
            const FOCUS_STABLE_TICKS: u64 = 20; // ~600ms at 30ms/tick
            const RGB_STABLE_TICKS: u64 = 70;
            let _ = wdl_hid_fallback_enabled(is_loq_15irx10);
            let _ = wdl_vendor_fallback_enabled(is_loq_15irx10);
            // Open HID only after we drop WinRT. Holding both at once is the flicker fight.
            let mut raw_hid: Option<(HidDevice, LampArrayReportIds, u16)> = None;
            let vendor_hids: Vec<VendorHidTarget> = Vec::new();
            log_to_file("Keepalive: WDL while focused; Windows hold + HID after click-away/minimize");

            loop {
                let (rgb_values, brightness, _speed, _effect_type, lamp_count) = {
                    let state = keepalive_state.lock().unwrap();
                    (
                        state.rgb_values,
                        state.brightness,
                        state.speed,
                        state.effect_type,
                        state.lamp_count,
                    )
                };

                let is_vivid_frame = is_vivid_rgb_frame(&rgb_values);
                if is_vivid_frame {
                    last_non_black_rgb = rgb_values;
                }
                // Paint the real frame, including dim/black steps. Substituting the
                // last bright color made Breath/Wave skip their dark phase.
                let paint_rgb = rgb_values;
                let hold_rgb = if is_vivid_frame {
                    rgb_values
                } else {
                    last_non_black_rgb
                };
                if hold_rgb != last_hold_rgb {
                    last_hold_rgb = hold_rgb;
                    rgb_unchanged_ticks = 0;
                } else {
                    rgb_unchanged_ticks = rgb_unchanged_ticks.saturating_add(1);
                }
                let is_static_frame = rgb_unchanged_ticks >= RGB_STABLE_TICKS;

                let window_is_active = keepalive_window_active.load(Ordering::Relaxed);
                if last_logged_focus != Some(window_is_active) {
                    log_to_file(&format!(
                        "GUI: window_active={} focused_ticks={} unfocused_ticks={} holding_lamp={}",
                        window_is_active, focused_ticks, unfocused_ticks, holding_lamp
                    ));
                    last_logged_focus = Some(window_is_active);
                }
                if window_is_active {
                    focused_ticks = focused_ticks.saturating_add(1);
                    unfocused_ticks = 0;
                } else {
                    unfocused_ticks = unfocused_ticks.saturating_add(1);
                    focused_ticks = 0;
                }

                if holding_lamp {
                    // Always release on click-away/minimize. Keeping exclusive LampArray
                    // while unfocused makes Windows ignore our writes, so Swipe/Disco die.
                    if unfocused_ticks >= FOCUS_STABLE_TICKS && is_vivid_rgb_frame(&hold_rgb) {
                        log_to_file(&format!(
                            "UNFOCUS: drop LampArray static_frame={} rgb=[{},{},{} {},{},{} {},{},{} {},{},{}]",
                            is_static_frame,
                            hold_rgb[0], hold_rgb[1], hold_rgb[2],
                            hold_rgb[3], hold_rgb[4], hold_rgb[5],
                            hold_rgb[6], hold_rgb[7], hold_rgb[8],
                            hold_rgb[9], hold_rgb[10], hold_rgb[11]
                        ));
                        set_controlled_by_foreground_app(false);
                        if let Ok(mut guard) = keepalive_lamp_array.lock() {
                            *guard = None;
                        }
                        holding_lamp = false;
                        pending_ambient_off = false;
                        let _ = reg_add_dword(r"HKCU\Software\Microsoft\Lighting", "AmbientLightingEnabled", 0);
                        raw_hid = None;
                        raw_hid = open_raw_lamp_array_hid(lamp_count);
                        log_to_file(&format!(
                            "UNFOCUS: HID LampArray feature reports {}",
                            if raw_hid.is_some() { "opened" } else { "unavailable" }
                        ));
                        log_fight_snapshot("after-drop", false, window_is_active, is_static_frame);
                    }
                } else if focused_ticks >= 3 {
                    log_to_file(&format!(
                        "FIGHT: reopen LampArray focused_ticks={} rgb_unchanged_ticks={}",
                        focused_ticks, rgb_unchanged_ticks
                    ));
                    raw_hid = None;
                    set_controlled_by_foreground_app(true);
                    if let Ok(mut guard) = keepalive_lamp_array.lock() {
                        if guard.is_none() {
                            *guard = open_wdl_lamp_array(&reopen_device_id);
                            if guard.is_some() {
                                log_to_file("WDL: re-acquired LampArray after focus");
                                pending_ambient_off = true;
                            }
                        }
                    }
                    holding_lamp = true;
                    log_fight_snapshot("after-reopen", true, window_is_active, is_static_frame);
                }

                let lc = lamp_count as usize;
                let zones = std::cmp::min(4, lc);
                let intensity = (brightness as f64 / 100.0).clamp(0.01, 1.0);

                // --- Path 1: WinRT direct color set only while we hold LampArray ---
                if holding_lamp {
                    if let Ok(guard) = keepalive_lamp_array.lock() {
                        if let Some(ref lamp_array) = *guard {
                            if zones > 0 {
                                for z in 0..zones {
                                    let zone_start = (z * lc) / zones;
                                    let zone_end = ((z + 1) * lc) / zones;
                                    if zone_start >= zone_end { continue; }
                                    let r = (paint_rgb[z * 3] as f64 * intensity) as u8;
                                    let g = (paint_rgb[z * 3 + 1] as f64 * intensity) as u8;
                                    let b = (paint_rgb[z * 3 + 2] as f64 * intensity) as u8;
                                    let color = Color { A: 255, R: r, G: g, B: b };
                                    let indices: Vec<i32> = (zone_start..zone_end).map(|i| i as i32).collect();
                                    if let Err(e) = lamp_array.SetSingleColorForIndices(color, &indices) {
                                        set_color_failures += 1;
                                        if reverse_engineering_mode_enabled() && set_color_failures % 200 == 1 {
                                            log_to_file(&format!(
                                                "RE-WDL [keepalive]: SetSingleColorForIndices failed count={} err={}",
                                                set_color_failures, e
                                            ));
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if pending_ambient_off {
                        let _ = reg_add_dword(r"HKCU\Software\Microsoft\Lighting", "AmbientLightingEnabled", 0);
                        for device_key in reg_query_subkeys(r"HKCU\Software\Microsoft\Lighting\Devices") {
                            let _ = reg_add_dword(&device_key, "AmbientLightingEnabled", 0);
                        }
                        pending_ambient_off = false;
                        log_to_file("WDL: Windows hold lighting off while this app is in control");
                    }
                }

                // Re-assert playlist activity periodically to survive background state changes.
                // OverrideZIndex must only be set before Start(); calling it after start fails.
                if tick % 20 == 0 {
                    if let Some(ref playlist) = keepalive_playlist {
                        if let Err(e) = playlist.Start() {
                            if reverse_engineering_mode_enabled() {
                                log_to_file(&format!(
                                    "RE-WDL [keepalive]: playlist Start failed err={}",
                                    e
                                ));
                            }
                        }
                    }
                }

                let callback_now = keepalive_callback_heartbeat.load(Ordering::Relaxed);
                if callback_now == last_callback_heartbeat {
                    callback_stall_ticks += 1;
                } else {
                    callback_stall_ticks = 0;
                    last_callback_heartbeat = callback_now;
                }

                // After click-away we do not hold WinRT. Paint the real 4-zone
                // frames through HID LampArray feature reports (host control).
                if !holding_lamp {
                    if let Some((ref dev, ref report_ids, raw_lc)) = raw_hid {
                        let intensity_byte = (intensity * 255.0) as u8;
                        write_lamp_array_zones(dev, report_ids, raw_lc, &paint_rgb, intensity_byte);
                        if report_ids.control != 0 && tick % 20 == 0 {
                            let _ = send_lamp_array_feature(dev, &[report_ids.control, 0x00]);
                        }
                    } else if tick % 60 == 0 {
                        raw_hid = open_raw_lamp_array_hid(lamp_count);
                    }
                }

                // --- Minimal logging every ~5s (only when raw HID is active) ---
                tick += 1;
                if tick % 800 == 0 {
                    log_fight_snapshot("periodic", holding_lamp, window_is_active, is_static_frame);
                    log_to_file(&format!(
                        "Keepalive: raw_hid={}, vendor_hids={}, gen7={}, window_active={}, holding_lamp={}, static_frame={}, tick={}, callback_heartbeat={}, callback_stall_ticks={}, set_color_failures={}, rgb=[{},{},{},{},{},{},{},{},{},{},{},{}]",
                        raw_hid.is_some(), vendor_hids.len(),
                        vendor_hids.iter().any(|target| target.gen7_controller.is_some()),
                        keepalive_window_active.load(Ordering::Relaxed),
                        holding_lamp,
                        is_static_frame,
                        tick,
                        callback_now,
                        callback_stall_ticks,
                        set_color_failures,
                        paint_rgb[0], paint_rgb[1], paint_rgb[2],
                        paint_rgb[3], paint_rgb[4], paint_rgb[5],
                        paint_rgb[6], paint_rgb[7], paint_rgb[8],
                        paint_rgb[9], paint_rgb[10], paint_rgb[11]
                    ));
                }

                // Sleep is reduced to 30ms so raw HID writes happen ~33 times/sec
                thread::sleep(Duration::from_millis(30));
            }
        });

        #[cfg(debug_assertions)]
        eprintln!("[DEBUG] Windows Dynamic Lighting: effect playlist started (persistent control)");

        let current_state = LightingState {
            effect_type: BaseEffects::Static,
            speed: 1,
            brightness: 50,
            rgb_values: [0; 12],
        };

        return Some(Keyboard {
            keyboard_hid: None,
            win_lamp_array: shared_lamp,
            win_color_state: Some(shared_state),
            _win_playlist: playlist,
            _win_keepalive: Some(keepalive),
            current_state,
            stop_signal: stop_signal.clone(),
            window_active,
            protocol: Protocol::WindowsDynamicLighting { lamp_count },
        });
    }

    None
}

pub fn get_keyboard(stop_signal: Arc<AtomicBool>) -> Result<Keyboard> {
    log_to_file("=== get_keyboard() starting ===");
    if reverse_engineering_mode_enabled() {
        log_to_file("RE-MODE enabled via LEGION_RGB_REVERSE_MODE");
    }
    let api: HidApi = HidApi::new()?;
    log_hid_inventory(&api, "get_keyboard");

    // Log ALL HID interfaces for VID 0x048d so we can see what's available
    for d in api.device_list() {
        if d.vendor_id() == 0x048d {
            log_to_file(&format!(
                "  HID: VID={:#06x} PID={:#06x} usage_page={:#06x} usage={:#06x} iface={} path={}",
                d.vendor_id(), d.product_id(), d.usage_page(), d.usage(),
                d.interface_number(),
                d.path().to_string_lossy()
            ));
        }
    }

    #[cfg(debug_assertions)]
    {
        eprintln!("[DEBUG] Searching for keyboard devices...");
        for d in api.device_list() {
            if d.vendor_id() == 0x048d {
                #[cfg(target_os = "windows")]
                eprintln!(
                    "[DEBUG]   VID: {:#06x}  PID: {:#06x}  Usage Page: {:#06x}  Usage: {:#06x}",
                    d.vendor_id(), d.product_id(), d.usage_page(), d.usage()
                );

                #[cfg(not(target_os = "windows"))]
                eprintln!(
                    "[DEBUG]   VID: {:#06x}  PID: {:#06x}",
                    d.vendor_id(),
                    d.product_id()
                );
            }
        }
    }

    #[cfg(target_os = "windows")]
    let is_loq_15irx10 = api.device_list().any(|d| d.vendor_id() == 0x048d && d.product_id() == 0xc693);

    // --- LOQ 15IRX10 dedicated flow ---
    #[cfg(target_os = "windows")]
    if is_loq_15irx10 {
        log_to_file("LOQ 15IRX10 detected: WDL while focused; Windows hold lighting after click-away");

        if should_disable_windows_ambient_lighting_for_loq() {
            disable_windows_dynamic_lighting_ambient();
        }

        // HID/vendor can open successfully on this model without driving the LEDs.
        // Keep Windows Dynamic Lighting first unless HID-first is explicitly requested.
        if loq_hid_first_enabled() {
            log_to_file("LOQ: LEGION_RGB_LOQ_HID_FIRST enabled, trying HID/vendor before Windows Dynamic Lighting");
            if let Some(keyboard) = try_lamp_array_keyboard(&api, &stop_signal) {
                log_to_file("SUCCESS: LOQ path using LampArray managed protocol");
                return Ok(keyboard);
            }
            if let Some(keyboard) = try_loq_vendor_keyboard(&api, &stop_signal) {
                log_to_file("SUCCESS: LOQ path using vendor legacy protocol fallback");
                return Ok(keyboard);
            }
        }

        // 1) Prefer Windows Dynamic Lighting. This is the path that actually updates LOQ 15IRX10 keys.
        if let Some(keyboard) = try_windows_dynamic_lighting(&stop_signal) {
            log_to_file("SUCCESS: LOQ path using Windows Dynamic Lighting API");
            return Ok(keyboard);
        }

        // 2) Managed LampArray path (with vendor keepalive in background)
        if let Some(keyboard) = try_lamp_array_keyboard(&api, &stop_signal) {
            log_to_file("SUCCESS: LOQ path using LampArray managed protocol");
            return Ok(keyboard);
        }

        // 3) Fallback directly to vendor endpoint legacy writes
        if let Some(keyboard) = try_loq_vendor_keyboard(&api, &stop_signal) {
            log_to_file("SUCCESS: LOQ path using vendor legacy protocol fallback");
            return Ok(keyboard);
        }

        log_to_file("LOQ dedicated flow failed; continuing with generic fallback strategy");
    }

    // --- Try raw HID LampArray first (direct HID, not subject to Windows DL service priority) ---
    #[cfg(target_os = "windows")]
    {
        log_to_file("Trying raw HID LampArray...");
        if let Some(keyboard) = try_lamp_array_keyboard(&api, &stop_signal) {
            log_to_file("SUCCESS: Using raw HID LampArray protocol");
            #[cfg(debug_assertions)]
            eprintln!("[DEBUG] Using raw HID LampArray protocol (not subject to Windows DL priority)");
            return Ok(keyboard);
        }
        log_to_file("HID LampArray not available, trying Windows Dynamic Lighting...");
        #[cfg(debug_assertions)]
        eprintln!("[DEBUG] HID LampArray not available, trying Windows Dynamic Lighting...");
    }

    // --- Try Windows Dynamic Lighting WinRT API (may be deprioritized when app is minimized) ---
    #[cfg(target_os = "windows")]
    {
        log_to_file("Trying Windows Dynamic Lighting WinRT API...");
        if let Some(keyboard) = try_windows_dynamic_lighting(&stop_signal) {
            log_to_file("SUCCESS: Using Windows Dynamic Lighting API");
            #[cfg(debug_assertions)]
            eprintln!("[DEBUG] Using Windows Dynamic Lighting API");
            return Ok(keyboard);
        }
        log_to_file("WDL not available, falling back to legacy protocol");
        #[cfg(debug_assertions)]
        eprintln!("[DEBUG] Windows Dynamic Lighting not available, falling back to legacy protocol");
    }

    // --- Strategy 1: exact 4-tuple match (most reliable) ---
    #[cfg(target_os = "windows")]
    let info = api.device_list().find(|d| {
        let t = (d.vendor_id(), d.product_id(), d.usage_page(), d.usage());
        KNOWN_DEVICE_INFOS.iter().any(|k| *k == t)
    });

    #[cfg(not(target_os = "windows"))]
    let info = api
        .device_list()
        .find(|d| is_known_pid(d.vendor_id(), d.product_id()));

    // --- Strategy 2: VID+PID match, prefer vendor-specific usage pages (>= 0xFF00) ---
    //     This handles newer models like the LOQ 15IRX10 whose LED-control
    //     interface may report a different usage page than 0xff89.
    #[cfg(target_os = "windows")]
    let info = info.or_else(|| {
        api.device_list().find(|d| {
            is_known_pid(d.vendor_id(), d.product_id()) && d.usage_page() >= 0xFF00
        })
    });

    // --- Strategy 3: VID+PID match, any HID interface (last resort) ---
    let info = info.or_else(|| {
        api.device_list().find(|d| is_known_pid(d.vendor_id(), d.product_id()))
    });

    let info = info.ok_or(error::Error::DeviceNotFound)?;

    #[cfg(debug_assertions)]
    #[cfg(target_os = "windows")]
    eprintln!(
        "[DEBUG] Opening device — VID: {:#06x}, PID: {:#06x}, Usage Page: {:#06x}, Usage: {:#06x}",
        info.vendor_id(), info.product_id(), info.usage_page(), info.usage()
    );

    #[cfg(debug_assertions)]
    #[cfg(not(target_os = "windows"))]
    eprintln!(
        "[DEBUG] Opening device — VID: {:#06x}, PID: {:#06x}",
        info.vendor_id(),
        info.product_id()
    );

    let keyboard_hid: HidDevice = info.open_device(&api)?;

    #[cfg(debug_assertions)]
    eprintln!("[DEBUG] Device opened successfully");

    let current_state: LightingState = LightingState {
        effect_type: BaseEffects::Static,
        speed: 1,
        brightness: 1,
        rgb_values: [0; 12],
    };

    let mut keyboard = Keyboard {
        keyboard_hid: Some(keyboard_hid),
        #[cfg(target_os = "windows")]
        win_lamp_array: Arc::new(std::sync::Mutex::new(None)),
        #[cfg(target_os = "windows")]
        win_color_state: None,
        #[cfg(target_os = "windows")]
        _win_playlist: None,
        #[cfg(target_os = "windows")]
        _win_keepalive: None,
        current_state,
        stop_signal,
        window_active: Arc::new(AtomicBool::new(true)),
        protocol: Protocol::Legacy(WriteMethod::Unknown),
    };

    keyboard.refresh()?;
    Ok(keyboard)
}

pub fn find_possible_keyboards() -> Result<Vec<String>> {
    let api: HidApi = HidApi::new()?;

    #[cfg(target_os = "windows")]
    let mut list = api
        .device_list()
        .filter(|d| d.vendor_id() == 0x048d)
        .map(|d| {
            format!(
                "{:#06x}:{:#06x} (usage_page: {:#06x}, usage: {:#06x})",
                d.vendor_id(),
                d.product_id(),
                d.usage_page(),
                d.usage()
            )
        })
        .collect::<Vec<String>>();

    #[cfg(not(target_os = "windows"))]
    let mut list = api
        .device_list()
        .filter(|d| d.vendor_id() == 0x048d)
        .map(|d| format!("{:#06x}:{:#06x}", d.vendor_id(), d.product_id()))
        .collect::<Vec<String>>();

    list.dedup();
    Ok(list)
}
