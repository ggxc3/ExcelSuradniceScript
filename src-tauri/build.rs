use std::{fs, path::Path};

const FALLBACK_ICON_BYTES: &[u8] = &[
    0, 0, 1, 0, 1, 0, 1, 1, 0, 0, 1, 0, 32, 0, 48, 0, 0, 0, 22, 0, 0, 0, 40, 0, 0, 0, 1, 0, 0, 0,
    2, 0, 0, 0, 1, 0, 32, 0, 0, 0, 0, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
    0, 0, 198, 104, 0, 255, 0, 0, 0, 0,
];

fn ensure_windows_icon_exists() {
    let icon_path = Path::new("icons/icon.ico");

    if icon_path.exists() {
        return;
    }

    if let Some(parent) = icon_path.parent() {
        fs::create_dir_all(parent).expect("Failed to create icons directory");
    }

    fs::write(icon_path, FALLBACK_ICON_BYTES).expect("Failed to write fallback icons/icon.ico");
    println!("cargo:warning=icons/icon.ico was missing; generated a fallback icon for tauri-build");
}

fn main() {
    ensure_windows_icon_exists();
    tauri_build::build();
}
