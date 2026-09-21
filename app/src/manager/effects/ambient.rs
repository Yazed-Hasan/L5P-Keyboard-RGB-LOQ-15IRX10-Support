use std::{
    sync::atomic::Ordering,
    thread,
    time::{Duration, Instant},
};

use fast_image_resize as fr;

use fr::Resizer;
use scrap::{Capturer, Display, Frame, TraitCapturer, TraitPixelBuffer};

use crate::manager::{effects::lamps, Inner};

#[derive(Clone, Copy)]
struct ScreenDimensions {
    src: (u32, u32),
    dest: (u32, u32),
}

pub fn play(manager: &mut Inner, fps: u8, saturation_boost: f32) {
    while !manager.stop_signals.manager_stop_signal.load(Ordering::SeqCst) {
        //Display setup
        let display = Display::all().unwrap().remove(0);

        let mut capturer = Capturer::new(display).expect("Couldn't begin capture.");

        let strip_w = manager.lamp_n().max(4) as u32;
        let dimensions = ScreenDimensions {
            src: (capturer.width() as u32, capturer.height() as u32),
            dest: (strip_w, 1),
        };

        let seconds_per_frame = Duration::from_nanos(1_000_000_000 / u64::from(fps));
        let mut resizer = fr::Resizer::new();

        #[cfg(target_os = "windows")]
        let mut try_gdi = 1;

        while !manager.stop_signals.keyboard_stop_signal.load(Ordering::SeqCst) {
            let now = Instant::now();

            #[allow(clippy::single_match)]
            match capturer.frame(seconds_per_frame) {
                Ok(frame) => {
                    let rgb = process_frame(frame, dimensions, &mut resizer, saturation_boost);
                    manager.paint_lamps(&rgb);
                    #[cfg(target_os = "windows")]
                    {
                        try_gdi = 0;
                    }
                }
                Err(error) => match error.kind() {
                    std::io::ErrorKind::WouldBlock =>
                    {
                        #[cfg(target_os = "windows")]
                        if try_gdi > 0 && !capturer.is_gdi() {
                            if try_gdi > 3 {
                                capturer.set_gdi();
                                try_gdi = 0;
                            }
                            try_gdi += 1;
                        }
                    }
                    _ =>
                    {
                        #[cfg(windows)]
                        if !capturer.is_gdi() {
                            capturer.set_gdi();
                            continue;
                        }
                    }
                },
            }

            let elapsed_time = now.elapsed();
            if elapsed_time < seconds_per_frame {
                thread::sleep(seconds_per_frame - elapsed_time);
            }
        }
    }
}

fn process_frame(frame: Frame, dimensions: ScreenDimensions, resizer: &mut Resizer, saturation_boost: f32) -> Vec<[u8; 3]> {
    // Adapted from https://github.com/Cykooz/fast_image_resize#resize-image
    let Frame::PixelBuffer(buf) = frame else {
        unreachable!("Attempted to extract vec from Texture variant in the Ambient effect");
    };

    let frame_vec = buf.data().to_vec();
    let src_image = fr::images::Image::from_vec_u8(dimensions.src.0, dimensions.src.1, frame_vec, fr::PixelType::U8x4).unwrap();
    let mut dst_image = fr::images::Image::new(dimensions.dest.0, dimensions.dest.1, fr::PixelType::U8x4);
    resizer.resize(&src_image, &mut dst_image, None).unwrap();

    let bgr_arr = dst_image.buffer();
    let mut rgba = vec![0u8; dimensions.dest.0 as usize * 4];
    for (src, dst) in bgr_arr.chunks_exact(4).zip(rgba.chunks_exact_mut(4)) {
        dst[0] = src[2];
        dst[1] = src[1];
        dst[2] = src[0];
        dst[3] = src[3];
    }

    let mut img = photon_rs::PhotonImage::new(rgba.to_vec(), dimensions.dest.0, 1);
    photon_rs::colour_spaces::saturate_hsv(&mut img, saturation_boost);

    let raw = img.get_raw_pixels();
    raw.chunks_exact(4)
        .map(|src| lamps::scale_rgb([src[0], src[1], src[2]], 1.0))
        .collect()
}
