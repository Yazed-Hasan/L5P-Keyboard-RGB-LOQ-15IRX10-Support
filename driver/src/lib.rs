use error::{RangeError, RangeErrorKind, Result};
use hidapi::{HidApi, HidDevice};
use std::{
    sync::{
        atomic::{AtomicBool, Ordering},
        Arc,
    },
    thread,
    time::Duration,
};

pub mod error;

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
    lamp_count: u16,
}

pub const SPEED_RANGE: std::ops::RangeInclusive<u8> = 1..=4;
pub const BRIGHTNESS_RANGE: std::ops::RangeInclusive<u8> = 1..=2;
pub const ZONE_RANGE: std::ops::RangeInclusive<u8> = 0..=3;

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
    win_lamp_array: Option<LampArray>,
    #[cfg(target_os = "windows")]
    win_color_state: Option<Arc<std::sync::Mutex<WinLampColors>>>,
    #[cfg(target_os = "windows")]
    _win_playlist: Option<LampArrayEffectPlaylist>,
    current_state: LightingState,
    stop_signal: Arc<AtomicBool>,
    protocol: Protocol,
}

#[allow(dead_code)]
impl Keyboard {
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
                }

                // Direct write as backup for when effect playlist is paused (app backgrounded)
                if let Some(ref lamp_array) = self.win_lamp_array {
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
        { matches!(self.protocol, Protocol::WindowsDynamicLighting { .. }) }
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
                    8 => report_id = value as u8, // Report ID
                    _ => {}
                }
            }
            2 => {
                // Local item — Usage (tag 0)
                if tag == 0 && usage_page == LAMP_ARRAY_USAGE_PAGE as u32 {
                    match value {
                        v if v == USAGE_LAMP_ARRAY_ATTRIBUTES_REPORT => {
                            ids.attributes = report_id;
                            found_any = true;
                        }
                        v if v == USAGE_LAMP_ATTRIBUTES_REQUEST_REPORT => {
                            ids.lamp_attr_request = report_id;
                            found_any = true;
                        }
                        v if v == USAGE_LAMP_ATTRIBUTES_RESPONSE_REPORT => {
                            ids.lamp_attr_response = report_id;
                            found_any = true;
                        }
                        v if v == USAGE_LAMP_MULTI_UPDATE_REPORT => {
                            ids.multi_update = report_id;
                            found_any = true;
                        }
                        v if v == USAGE_LAMP_RANGE_UPDATE_REPORT => {
                            ids.range_update = report_id;
                            found_any = true;
                        }
                        v if v == USAGE_LAMP_ARRAY_CONTROL_REPORT => {
                            ids.control = report_id;
                            found_any = true;
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }

        i += 1 + size;
    }

    if found_any {
        Some(ids)
    } else {
        None
    }
}

/// Try to open a HID LampArray interface for a known device.
/// Returns a Keyboard configured for the LampArray protocol, or None.
#[cfg(target_os = "windows")]
fn try_lamp_array_keyboard(api: &HidApi, stop_signal: &Arc<AtomicBool>) -> Option<Keyboard> {
    let lamp_info = api.device_list().find(|d| {
        is_known_pid(d.vendor_id(), d.product_id())
            && d.usage_page() == LAMP_ARRAY_USAGE_PAGE
            && d.usage() == LAMP_ARRAY_USAGE
    })?;

    #[cfg(debug_assertions)]
    eprintln!(
        "[DEBUG] Found LampArray interface — VID: {:#06x}, PID: {:#06x}, Usage Page: {:#06x}",
        lamp_info.vendor_id(),
        lamp_info.product_id(),
        lamp_info.usage_page()
    );

    let device = lamp_info.open_device(api).ok()?;

    // Parse the report descriptor to discover LampArray report IDs
    let mut desc_buf = [0u8; 4096];
    let desc_len = device.get_report_descriptor(&mut desc_buf).ok()?;
    let report_ids = parse_lamp_array_report_ids(&desc_buf[..desc_len])?;

    #[cfg(debug_assertions)]
    eprintln!(
        "[DEBUG] LampArray report IDs — attrs:{}, req:{}, resp:{}, multi:{}, range:{}, ctrl:{}",
        report_ids.attributes,
        report_ids.lamp_attr_request,
        report_ids.lamp_attr_response,
        report_ids.multi_update,
        report_ids.range_update,
        report_ids.control
    );

    // Read LampArrayAttributes to get lamp count
    let mut attr_buf = [0u8; 64];
    attr_buf[0] = report_ids.attributes;
    let n = device.get_feature_report(&mut attr_buf).ok()?;
    if n < 3 {
        return None;
    }
    let lamp_count = u16::from_le_bytes([attr_buf[1], attr_buf[2]]);
    if lamp_count == 0 {
        return None;
    }

    #[cfg(debug_assertions)]
    eprintln!("[DEBUG] LampArray: {} lamps detected", lamp_count);

    // Take host control (autonomous_mode = 0)
    let control_buf = [report_ids.control, 0x00];
    device.send_feature_report(&control_buf).ok()?;

    #[cfg(debug_assertions)]
    eprintln!("[DEBUG] LampArray: host control acquired");

    let current_state = LightingState {
        effect_type: BaseEffects::Static,
        speed: 1,
        brightness: 1,
        rgb_values: [0; 12],
    };

    Some(Keyboard {
        keyboard_hid: Some(device),
        #[cfg(target_os = "windows")]
        win_lamp_array: None,
        #[cfg(target_os = "windows")]
        win_color_state: None,
        #[cfg(target_os = "windows")]
        _win_playlist: None,
        current_state,
        stop_signal: stop_signal.clone(),
        protocol: Protocol::LampArrayHid { report_ids, lamp_count },
    })
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
        let id_str = id.to_string().to_lowercase();

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

        #[cfg(debug_assertions)]
        eprintln!(
            "[DEBUG] Windows Dynamic Lighting: opened LampArray with {} lamps",
            lamp_count
        );

        // Shared color state: the effect callback reads, refresh() writes
        let shared_state = Arc::new(std::sync::Mutex::new(WinLampColors {
            rgb_values: [0; 12],
            brightness: 50,
            lamp_count,
        }));

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
        let _ = effect.UpdateRequested(&TypedEventHandler::new(
            move |_effect, args: &Option<LampArrayUpdateRequestedEventArgs>| {
                if let Some(args) = args {
                    let state = shared_clone.lock().unwrap();
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

        // Start a persistent effect playlist — prevents Windows from overriding colors
        let playlist = match LampArrayEffectPlaylist::new() {
            Ok(p) => p,
            Err(_) => continue,
        };
        let _ = playlist.Append(&effect);
        let _ = playlist.SetRepetitionMode(LampArrayRepetitionMode::Forever);
        let _ = playlist.Start();

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
            win_lamp_array: Some(lamp_array),
            win_color_state: Some(shared_state),
            _win_playlist: Some(playlist),
            current_state,
            stop_signal: stop_signal.clone(),
            protocol: Protocol::WindowsDynamicLighting { lamp_count },
        });
    }

    None
}

pub fn get_keyboard(stop_signal: Arc<AtomicBool>) -> Result<Keyboard> {
    let api: HidApi = HidApi::new()?;

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

    // --- Try Windows Dynamic Lighting WinRT API first (most reliable for LOQ 15IRX10) ---
    #[cfg(target_os = "windows")]
    {
        if let Some(keyboard) = try_windows_dynamic_lighting(&stop_signal) {
            #[cfg(debug_assertions)]
            eprintln!("[DEBUG] Using Windows Dynamic Lighting API");
            return Ok(keyboard);
        }
        #[cfg(debug_assertions)]
        eprintln!("[DEBUG] Windows Dynamic Lighting not available, trying HID LampArray...");
    }

    // --- Try raw HID LampArray (direct HID, may conflict with Windows DL service) ---
    #[cfg(target_os = "windows")]
    {
        if let Some(keyboard) = try_lamp_array_keyboard(&api, &stop_signal) {
            #[cfg(debug_assertions)]
            eprintln!("[DEBUG] Using raw HID LampArray protocol");
            return Ok(keyboard);
        }
        #[cfg(debug_assertions)]
        eprintln!("[DEBUG] HID LampArray not available, falling back to legacy protocol");
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
        win_lamp_array: None,
        #[cfg(target_os = "windows")]
        win_color_state: None,
        #[cfg(target_os = "windows")]
        _win_playlist: None,
        current_state,
        stop_signal,
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
