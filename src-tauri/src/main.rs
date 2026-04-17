#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use excel_suradnice_script::{run, ProcessRequest};

#[tauri::command]
fn process_workbook(request: ProcessRequest) -> Result<String, String> {
    run(request).map_err(|e| e.to_string())
}

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![process_workbook])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
