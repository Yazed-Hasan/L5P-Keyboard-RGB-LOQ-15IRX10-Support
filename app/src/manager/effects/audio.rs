use std::{
    f32::consts::PI,
    sync::{
        atomic::Ordering,
        Arc, Mutex,
    },
    thread,
    time::{Duration, Instant},
};

use crate::{
    enums::{AudioColorMode, AudioStyle, Effects},
    manager::{profile::Profile, Inner},
};

const FFT_SIZE: usize = 2048;
const SAMPLE_RATE: u32 = 48_000;

#[derive(Clone, Copy, Debug)]
pub struct AudioReactParams {
    pub sensitivity: f32,
    pub smoothness: f32,
    pub min_brightness: u8,
    pub bass: f32,
    pub mid: f32,
    pub treble: f32,
    pub presence: f32,
    pub squelch: f32,
    pub punch: f32,
    pub color_mode: AudioColorMode,
    pub style: AudioStyle,
}

impl AudioReactParams {
    pub fn from_effect(effect: Effects) -> Self {
        match effect {
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
            } => Self {
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
            }
            .normalized(),
            _ => Self::default(),
        }
    }

    pub fn normalized(self) -> Self {
        let or_one = |v: f32| if v <= 0.001 { 1.0 } else { v };
        Self {
            sensitivity: if self.sensitivity <= 0.05 {
                1.2
            } else {
                self.sensitivity.clamp(0.2, 5.0)
            },
            smoothness: self.smoothness.clamp(0.0, 0.95),
            min_brightness: self.min_brightness.min(80),
            bass: or_one(self.bass).clamp(0.0, 2.5),
            mid: or_one(self.mid).clamp(0.0, 2.5),
            treble: or_one(self.treble).clamp(0.0, 2.5),
            presence: or_one(self.presence).clamp(0.0, 2.5),
            squelch: self.squelch.clamp(0.0, 0.35),
            punch: self.punch.clamp(0.0, 2.0),
            color_mode: self.color_mode,
            style: self.style,
        }
    }
}

impl Default for AudioReactParams {
    fn default() -> Self {
        Self {
            sensitivity: 1.2,
            smoothness: 0.72,
            min_brightness: 6,
            bass: 1.0,
            mid: 1.0,
            treble: 1.0,
            presence: 1.0,
            squelch: 0.07,
            punch: 0.65,
            color_mode: AudioColorMode::Spectrum,
            style: AudioStyle::Levels,
        }
    }
}

pub fn play(manager: &mut Inner, profile: &Profile) {
    let params = manager.audio_params.clone();
    {
        let mut guard = params.lock().unwrap();
        *guard = AudioReactParams::from_effect(profile.effect);
    }

    let samples = Arc::new(Mutex::new(vec![0.0f32; FFT_SIZE]));
    let capture_ok = Arc::new(std::sync::atomic::AtomicBool::new(false));
    let stop = manager.stop_signals.manager_stop_signal.clone();
    let capture_stop = stop.clone();
    let capture_samples = samples.clone();
    let capture_flag = capture_ok.clone();

    let capture = thread::spawn(move || {
        #[cfg(target_os = "windows")]
        {
            if let Err(err) = capture_loopback(capture_samples, capture_stop, capture_flag) {
                legion_rgb_driver::debug_log(&format!("AUDIO: loopback failed: {err}"));
            }
        }
        #[cfg(not(target_os = "windows"))]
        {
            let _ = (capture_samples, capture_stop, capture_flag);
            legion_rgb_driver::debug_log("AUDIO: loopback is only available on Windows");
        }
    });

    legion_rgb_driver::debug_log("EFFECT: audio react (system playback)");

    let mut short_avg = [0.0f32; 4];
    let mut long_avg = [0.16f32; 4];
    let mut prev_short = [0.0f32; 4];
    let mut levels = [0.0f32; 4];
    let mut wave_phase = 0.0f32;
    let mut quiet_frames = 0u32;
    let mut last_tick = Instant::now();
    let mut last_log = Instant::now();
    let profile_rgb = profile.rgb_array();

    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        let now = Instant::now();
        let dt = now.saturating_duration_since(last_tick).as_secs_f32().clamp(0.008, 0.05);
        last_tick = now;
        let p = params.lock().unwrap().normalized();
        let frame = {
            let guard = samples.lock().unwrap();
            analyze_bands(&guard, SAMPLE_RATE)
        };

        let boosted = [
            frame[0] * p.bass,
            frame[1] * p.mid,
            frame[2] * p.treble,
            frame[3] * p.presence,
        ];
        let instant = boosted.iter().copied().fold(0.0f32, f32::max);
        if instant < p.squelch.max(0.02) {
            quiet_frames = quiet_frames.saturating_add(1);
        } else {
            quiet_frames = 0;
        }

        let short_tau = 0.045 + p.smoothness * 0.03;
        let long_tau = if quiet_frames > 90 { 2.8 } else { 1.15 };
        let attack_tau = 0.04 + p.smoothness * 0.05;
        let release_tau = 0.16 + p.smoothness * 0.55;

        for i in 0..4 {
            let raw = boosted[i].max(0.0);
            short_avg[i] = ema_toward(short_avg[i], raw, dt, short_tau);
            if quiet_frames <= 90 {
                long_avg[i] = ema_toward(long_avg[i], raw.max(0.04), dt, long_tau).max(0.06);
            }
            let rel = short_avg[i] / long_avg[i];
            let flux = (short_avg[i] - prev_short[i]).max(0.0) / long_avg[i];
            prev_short[i] = short_avg[i];
            let driven = (rel - 0.72).max(0.0) * (0.85 + p.sensitivity * 0.55) + flux * p.punch * 1.15;
            let target = if quiet_frames > 45 {
                0.0
            } else {
                driven.clamp(0.0, 1.0)
            };
            let tau = if target > levels[i] { attack_tau } else { release_tau };
            levels[i] = ema_toward(levels[i], target, dt, tau);
        }

        let rgb = render_audio(p, &levels, &profile_rgb, &mut wave_phase, dt);
        let _ = manager.keyboard.set_colors_to(&rgb);

        if last_log.elapsed() > Duration::from_secs(3) {
            legion_rgb_driver::debug_log(&format!(
                "AUDIO: capture={} bass={:.2} mid={:.2} treble={:.2} presence={:.2}",
                capture_ok.load(Ordering::Relaxed),
                levels[0],
                levels[1],
                levels[2],
                levels[3]
            ));
            last_log = Instant::now();
        }

        thread::sleep(Duration::from_millis(16));
    }

    let _ = capture.join();
}

fn ema_toward(current: f32, target: f32, dt: f32, tau: f32) -> f32 {
    let alpha = 1.0 - (-dt / tau.max(0.008)).exp();
    current + (target - current) * alpha.clamp(0.0, 1.0)
}

fn led_curve(energy: f32) -> f32 {
    energy.clamp(0.0, 1.0).powf(1.35)
}

fn render_audio(
    params: AudioReactParams,
    levels: &[f32; 4],
    profile_rgb: &[u8; 12],
    wave_phase: &mut f32,
    dt: f32,
) -> [u8; 12] {
    let floor = params.min_brightness as f32 / 100.0;
    if matches!(params.color_mode, AudioColorMode::Rainbow) && !matches!(params.style, AudioStyle::Wave) {
        *wave_phase = (*wave_phase + dt * (0.45 + levels[0] * 1.1)) % (2.0 * PI);
    }
    let mut energies = *levels;
    match params.style {
        AudioStyle::Levels => {
            energies = spatial_blur(levels);
        }
        AudioStyle::Pulse => {
            let pulse = (levels.iter().sum::<f32>() / 4.0).max(levels[0] * 0.85);
            energies = [pulse; 4];
        }
        AudioStyle::Wave => {
            *wave_phase = (*wave_phase + dt * (1.6 + levels[0] * 4.2)) % (2.0 * PI);
            let drive = levels[0].max(levels[1]);
            for (i, energy) in energies.iter_mut().enumerate() {
                let travel = 0.5 + 0.5 * ((*wave_phase) - i as f32 * 0.85).sin();
                *energy = (*energy * 0.42 + travel * drive * 0.58).clamp(0.0, 1.0);
            }
        }
        AudioStyle::Bloom => {
            energies[0] = levels[0];
            energies[1] = levels[1].max(levels[0] * 0.62);
            energies[2] = levels[2].max(levels[0] * 0.32 + levels[1] * 0.45);
            energies[3] = levels[3].max(levels[1] * 0.28 + levels[2] * 0.40);
        }
        AudioStyle::Center => {
            energies[0] = levels[0].max(levels[3] * 0.35);
            energies[1] = levels[1].max(levels[2] * 0.55);
            energies[2] = levels[2].max(levels[1] * 0.55);
            energies[3] = levels[3].max(levels[0] * 0.50);
        }
    }

    let mut rgb = [0u8; 12];
    for z in 0..4 {
        let amount = (floor + (1.0 - floor) * led_curve(energies[z])).clamp(0.0, 1.0);
        let (r, g, b) = match params.color_mode {
            AudioColorMode::Profile => (
                profile_rgb[z * 3],
                profile_rgb[z * 3 + 1],
                profile_rgb[z * 3 + 2],
            ),
            AudioColorMode::Spectrum => {
                const SPECTRUM: [[u8; 3]; 4] = [[255, 24, 48], [255, 140, 16], [36, 220, 120], [72, 120, 255]];
                (SPECTRUM[z][0], SPECTRUM[z][1], SPECTRUM[z][2])
            }
            AudioColorMode::Rainbow => {
                let hue = (z as f32 * 90.0 + *wave_phase * 12.0) % 360.0;
                hsv_to_rgb(hue, 0.92, 1.0)
            }
        };
        rgb[z * 3] = (r as f32 * amount) as u8;
        rgb[z * 3 + 1] = (g as f32 * amount) as u8;
        rgb[z * 3 + 2] = (b as f32 * amount) as u8;
    }
    rgb
}

fn spatial_blur(levels: &[f32; 4]) -> [f32; 4] {
    [
        levels[0] * 0.78 + levels[1] * 0.22,
        levels[1] * 0.62 + levels[0] * 0.19 + levels[2] * 0.19,
        levels[2] * 0.62 + levels[1] * 0.19 + levels[3] * 0.19,
        levels[3] * 0.78 + levels[2] * 0.22,
    ]
}

fn analyze_bands(samples: &[f32], sample_rate: u32) -> [f32; 4] {
    if samples.len() < FFT_SIZE {
        return [0.0; 4];
    }
    let mut re = [0.0f32; FFT_SIZE];
    let mut im = [0.0f32; FFT_SIZE];
    let n = FFT_SIZE as f32;
    for i in 0..FFT_SIZE {
        let t = 2.0 * PI * i as f32 / n;
        let window = 0.35875 - 0.48829 * t.cos() + 0.14128 * (2.0 * t).cos() - 0.01168 * (3.0 * t).cos();
        re[i] = samples[i] * window;
    }
    fft(&mut re, &mut im);

    let filters = [
        (25.0, 70.0, 200.0),
        (90.0, 320.0, 800.0),
        (400.0, 1400.0, 3500.0),
        (1800.0, 5000.0, 12000.0),
    ];
    let bin_hz = sample_rate as f32 / n;
    let mut bands = [0.0f32; 4];
    for (band, (lo, mid, hi)) in filters.iter().enumerate() {
        let mut acc = 0.0;
        let mut weight = 0.0;
        let start = (*lo / bin_hz).floor().max(1.0) as usize;
        let end = (*hi / bin_hz).ceil().min((FFT_SIZE / 2 - 1) as f32) as usize;
        for i in start..=end {
            let freq = i as f32 * bin_hz;
            let tri = if freq <= *mid {
                (freq - lo) / (mid - lo)
            } else {
                (hi - freq) / (hi - mid)
            }
            .clamp(0.0, 1.0);
            if tri <= 0.0 {
                continue;
            }
            let mag = (re[i] * re[i] + im[i] * im[i]).sqrt();
            let pink = (freq / 800.0).sqrt().clamp(0.35, 2.6);
            acc += mag * tri * pink;
            weight += tri;
        }
        let mean = if weight > 0.0 { acc / weight } else { 0.0 };
        bands[band] = (mean / 48.0).ln_1p();
    }
    bands
}

fn fft(re: &mut [f32], im: &mut [f32]) {
    let n = re.len();
    let mut j = 0usize;
    for i in 1..n {
        let mut bit = n >> 1;
        while j & bit != 0 {
            j ^= bit;
            bit >>= 1;
        }
        j ^= bit;
        if i < j {
            re.swap(i, j);
            im.swap(i, j);
        }
    }

    let mut len = 2;
    while len <= n {
        let ang = -2.0 * PI / len as f32;
        let (wlen_im, wlen_re) = ang.sin_cos();
        let half = len / 2;
        for i in (0..n).step_by(len) {
            let mut wr = 1.0;
            let mut wi = 0.0;
            for k in 0..half {
                let u_re = re[i + k];
                let u_im = im[i + k];
                let v_re = re[i + k + half] * wr - im[i + k + half] * wi;
                let v_im = re[i + k + half] * wi + im[i + k + half] * wr;
                re[i + k] = u_re + v_re;
                im[i + k] = u_im + v_im;
                re[i + k + half] = u_re - v_re;
                im[i + k + half] = u_im - v_im;
                let next_wr = wr * wlen_re - wi * wlen_im;
                wi = wr * wlen_im + wi * wlen_re;
                wr = next_wr;
            }
        }
        len <<= 1;
    }
}

fn hsv_to_rgb(h: f32, s: f32, v: f32) -> (u8, u8, u8) {
    let c = v * s;
    let x = c * (1.0 - ((h / 60.0) % 2.0 - 1.0).abs());
    let m = v - c;
    let (r, g, b) = match h {
        h if h < 60.0 => (c, x, 0.0),
        h if h < 120.0 => (x, c, 0.0),
        h if h < 180.0 => (0.0, c, x),
        h if h < 240.0 => (0.0, x, c),
        h if h < 300.0 => (x, 0.0, c),
        _ => (c, 0.0, x),
    };
    (((r + m) * 255.0) as u8, ((g + m) * 255.0) as u8, ((b + m) * 255.0) as u8)
}

#[cfg(target_os = "windows")]
fn capture_loopback(
    samples: Arc<Mutex<Vec<f32>>>,
    stop: Arc<std::sync::atomic::AtomicBool>,
    capture_ok: Arc<std::sync::atomic::AtomicBool>,
) -> Result<(), String> {
    use std::{
        collections::{HashMap, HashSet},
        sync::atomic::AtomicUsize,
    };

    wasapi::initialize_mta()
        .ok()
        .map_err(|_| "COM init failed".to_string())?;

    let per_device: Arc<Mutex<HashMap<String, Vec<f32>>>> = Arc::new(Mutex::new(HashMap::new()));
    let wanted: Arc<Mutex<HashSet<String>>> = Arc::new(Mutex::new(HashSet::new()));
    let live = Arc::new(AtomicUsize::new(0));
    let mut workers: HashMap<String, thread::JoinHandle<()>> = HashMap::new();
    let mut last_ids = HashSet::new();

    while !stop.load(Ordering::SeqCst) {
        let targets = list_render_targets();
        let ids: HashSet<String> = targets.iter().map(|(id, _)| id.clone()).collect();
        if let Ok(mut guard) = wanted.lock() {
            *guard = ids.clone();
        }

        if ids != last_ids {
            let names: Vec<String> = targets.iter().map(|(_, name)| name.clone()).collect();
            legion_rgb_driver::debug_log(&format!(
                "AUDIO: follow outputs [{}]",
                names.join(" | ")
            ));
            last_ids = ids.clone();
        }

        for (id, name) in targets {
            if workers.contains_key(&id) {
                continue;
            }
            let samples = samples.clone();
            let per_device = per_device.clone();
            let wanted = wanted.clone();
            let stop = stop.clone();
            let live = live.clone();
            let capture_ok = capture_ok.clone();
            workers.insert(
                id.clone(),
                thread::spawn(move || {
                    if let Err(err) = capture_one_output(&id, &name, samples, per_device, wanted, stop, live, capture_ok)
                    {
                        legion_rgb_driver::debug_log(&format!("AUDIO: {name} loopback ended: {err}"));
                    }
                }),
            );
        }

        let finished: Vec<String> = workers
            .iter()
            .filter(|(_, handle)| handle.is_finished())
            .map(|(id, _)| id.clone())
            .collect();
        for id in finished {
            if let Some(handle) = workers.remove(&id) {
                let _ = handle.join();
            }
        }
        capture_ok.store(live.load(Ordering::Relaxed) > 0, Ordering::Relaxed);
        thread::sleep(Duration::from_millis(350));
    }

    if let Ok(mut guard) = wanted.lock() {
        guard.clear();
    }
    for (_, handle) in workers {
        let _ = handle.join();
    }
    capture_ok.store(false, Ordering::Relaxed);
    Ok(())
}

#[cfg(target_os = "windows")]
fn list_render_targets() -> Vec<(String, String)> {
    use std::collections::BTreeMap;
    use wasapi::{get_default_device_for_role, DeviceCollection, Direction, Role};

    let mut found: BTreeMap<String, String> = BTreeMap::new();
    if let Ok(collection) = DeviceCollection::new(&Direction::Render) {
        if let Ok(count) = collection.get_nbr_devices() {
            for idx in 0..count {
                if let Ok(device) = collection.get_device_at_index(idx) {
                    if let (Ok(id), Ok(name)) = (device.get_id(), device.get_friendlyname()) {
                        found.insert(id, name);
                    }
                }
            }
        }
    }
    for role in [Role::Console, Role::Multimedia, Role::Communications] {
        if let Ok(device) = get_default_device_for_role(&Direction::Render, &role) {
            if let (Ok(id), Ok(name)) = (device.get_id(), device.get_friendlyname()) {
                found.entry(id).or_insert(name);
            }
        }
    }
    found.into_iter().collect()
}

#[cfg(target_os = "windows")]
fn publish_mix(per_device: &Mutex<std::collections::HashMap<String, Vec<f32>>>, samples: &Mutex<Vec<f32>>) {
    let mix = {
        let Ok(map) = per_device.lock() else {
            return;
        };
        if map.is_empty() {
            vec![0.0f32; FFT_SIZE]
        } else {
            let mut best: Option<(&Vec<f32>, f32)> = None;
            for buf in map.values() {
                let energy = buf.iter().map(|s| s * s).sum::<f32>();
                if best.as_ref().map(|(_, e)| energy > *e).unwrap_or(true) {
                    best = Some((buf, energy));
                }
            }
            best.map(|(buf, _)| buf.clone()).unwrap_or_else(|| vec![0.0f32; FFT_SIZE])
        }
    };
    if let Ok(mut guard) = samples.lock() {
        *guard = mix;
    }
}

#[cfg(target_os = "windows")]
fn capture_one_output(
    id: &str,
    name: &str,
    samples: Arc<Mutex<Vec<f32>>>,
    per_device: Arc<Mutex<std::collections::HashMap<String, Vec<f32>>>>,
    wanted: Arc<Mutex<std::collections::HashSet<String>>>,
    stop: Arc<std::sync::atomic::AtomicBool>,
    live: Arc<std::sync::atomic::AtomicUsize>,
    capture_ok: Arc<std::sync::atomic::AtomicBool>,
) -> Result<(), String> {
    use std::collections::VecDeque;
    use wasapi::{DeviceCollection, Direction, SampleType, StreamMode, WaveFormat};

    wasapi::initialize_mta()
        .ok()
        .map_err(|_| "COM init failed".to_string())?;

    let collection = DeviceCollection::new(&Direction::Render).map_err(|e| e.to_string())?;
    let count = collection.get_nbr_devices().map_err(|e| e.to_string())?;
    let mut device = None;
    for idx in 0..count {
        if let Ok(candidate) = collection.get_device_at_index(idx) {
            if candidate.get_id().ok().as_deref() == Some(id) {
                device = Some(candidate);
                break;
            }
        }
    }
    let device = device.ok_or_else(|| format!("output gone: {name}"))?;
    let mut audio_client = device.get_iaudioclient().map_err(|e| e.to_string())?;
    let format = WaveFormat::new(32, 32, &SampleType::Float, SAMPLE_RATE as usize, 2, None);
    let mode = StreamMode::EventsShared {
        autoconvert: true,
        buffer_duration_hns: 200_000,
    };
    audio_client
        .initialize_client(&format, &Direction::Capture, &mode)
        .map_err(|e| e.to_string())?;
    let event = audio_client.set_get_eventhandle().map_err(|e| e.to_string())?;
    let capture_client = audio_client.get_audiocaptureclient().map_err(|e| e.to_string())?;
    audio_client.start_stream().map_err(|e| e.to_string())?;
    live.fetch_add(1, Ordering::Relaxed);
    capture_ok.store(true, Ordering::Relaxed);
    legion_rgb_driver::debug_log(&format!("AUDIO: WASAPI loopback started on {name}"));

    let mut queue: VecDeque<u8> = VecDeque::new();
    let mut ring = vec![0.0f32; FFT_SIZE];
    let mut write_at = 0usize;
    let channels = 2usize;
    let bytes_per_sample = 4usize;
    let mut stale_reads = 0u8;

    while !stop.load(Ordering::SeqCst) {
        if let Ok(guard) = wanted.lock() {
            if !guard.contains(id) {
                break;
            }
        }
        if event.wait_for_event(80).is_err() {
            continue;
        }
        if capture_client.read_from_device_to_deque(&mut queue).is_err() {
            stale_reads = stale_reads.saturating_add(1);
            if stale_reads >= 8 {
                break;
            }
            continue;
        }
        stale_reads = 0;
        let frame_bytes = bytes_per_sample * channels;
        while queue.len() >= frame_bytes {
            let mut left = [0u8; 4];
            let mut right = [0u8; 4];
            for b in &mut left {
                *b = queue.pop_front().unwrap();
            }
            for b in &mut right {
                *b = queue.pop_front().unwrap();
            }
            let l = f32::from_le_bytes(left);
            let r = f32::from_le_bytes(right);
            ring[write_at] = (l + r) * 0.5;
            write_at = (write_at + 1) % FFT_SIZE;
        }
        let window = {
            let mut window = vec![0.0f32; FFT_SIZE];
            let (tail, head) = ring.split_at(write_at);
            let split = FFT_SIZE - write_at;
            window[split..].copy_from_slice(tail);
            window[..split].copy_from_slice(head);
            window
        };
        if let Ok(mut map) = per_device.lock() {
            map.insert(id.to_string(), window);
        }
        publish_mix(&per_device, &samples);
    }

    let _ = audio_client.stop_stream();
    if let Ok(mut map) = per_device.lock() {
        map.remove(id);
    }
    publish_mix(&per_device, &samples);
    live.fetch_sub(1, Ordering::Relaxed);
    capture_ok.store(live.load(Ordering::Relaxed) > 0, Ordering::Relaxed);
    Ok(())
}
