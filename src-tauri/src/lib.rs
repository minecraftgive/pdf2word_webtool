pub mod convert;

use serde::Serialize;
use tauri::{AppHandle, Emitter};

#[derive(Serialize, Clone)]
struct ProgressEvent {
    step: String,
    detail: String,
}

/// Convert PDF bytes to a .docx, emitting progress events along the way.
#[tauri::command]
fn convert_pdf(data: Vec<u8>, app: AppHandle) -> Result<Vec<u8>, String> {
    convert::pdf_to_docx(&data, |step, detail| {
        let _ = app.emit(
            "progress",
            ProgressEvent {
                step: step.to_string(),
                detail: detail.to_string(),
            },
        );
    })
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![convert_pdf])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
