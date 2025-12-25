use tauri::plugin::{Builder, TauriPlugin};
use tauri::Runtime;

#[tauri::command]
async fn compare_documents(
    source1_path: String,
    source2_path: String,
    output_path: String,
    _settings: Option<serde_json::Value>,
) -> Result<serde_json::Value, String> {
    let _older = tokio::fs::read(&source1_path)
        .await
        .map_err(|e| e.to_string())?;
    let _newer = tokio::fs::read(&source2_path)
        .await
        .map_err(|e| e.to_string())?;

    todo!("Tauri plugin not yet implemented - Phase 6");

    #[allow(unreachable_code)]
    Ok(serde_json::json!({
        "output_path": output_path,
        "status": "success"
    }))
}

pub fn init<R: Runtime>() -> TauriPlugin<R> {
    Builder::new("redline")
        .invoke_handler(tauri::generate_handler![compare_documents])
        .build()
}
