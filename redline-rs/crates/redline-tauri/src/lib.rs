use tauri::plugin::{Builder, TauriPlugin};
use tauri::Runtime;
use redline_core::{
    WmlComparer, WmlDocument, WmlComparerSettings,
    sml::{SmlComparer, SmlDocument, SmlComparerSettings},
    pml::{PmlComparer, PmlDocument, PmlComparerSettings},
};
use std::path::Path;

#[derive(serde::Serialize)]
struct ComparisonResponse {
    stats: serde_json::Value,
    output_path: String,
}

#[tauri::command]
async fn compare_documents(
    source1_path: String,
    source2_path: String,
    output_path: String,
    settings: Option<serde_json::Value>,
) -> Result<ComparisonResponse, String> {
    let path1 = Path::new(&source1_path);
    let ext = path1.extension().and_then(|e| e.to_str()).unwrap_or("").to_lowercase();
    
    // Read files
    let older_bytes = tokio::fs::read(&source1_path).await.map_err(|e| e.to_string())?;
    let newer_bytes = tokio::fs::read(&source2_path).await.map_err(|e| e.to_string())?;

    if ext == "docx" || ext == "docm" {
        let older = WmlDocument::from_bytes(&older_bytes).map_err(|e| e.to_string())?;
        let newer = WmlDocument::from_bytes(&newer_bytes).map_err(|e| e.to_string())?;
        
        let wml_settings: WmlComparerSettings = if let Some(s) = settings {
            serde_json::from_value(s).unwrap_or_default()
        } else {
            WmlComparerSettings::default()
        };
        
        let result = WmlComparer::compare(&older, &newer, Some(&wml_settings)).map_err(|e| e.to_string())?;
        
        // Save output
        tokio::fs::write(&output_path, &result.document).await.map_err(|e| e.to_string())?;
        
        return Ok(ComparisonResponse {
            stats: serde_json::json!({
                "revision_count": result.revision_count,
                "insertions": result.insertions,
                "deletions": result.deletions,
            }),
            output_path,
        });
    } else if ext == "xlsx" || ext == "xlsm" {
        let older = SmlDocument::from_bytes(&older_bytes).map_err(|e| e.to_string())?;
        let newer = SmlDocument::from_bytes(&newer_bytes).map_err(|e| e.to_string())?;
        
        let sml_settings: SmlComparerSettings = if let Some(s) = settings {
            serde_json::from_value(s).unwrap_or_default()
        } else {
            SmlComparerSettings::default()
        };
        
        let (marked_doc, result) = SmlComparer::compare_and_render(&older, &newer, Some(&sml_settings)).map_err(|e| e.to_string())?;
        let output_bytes = marked_doc.to_bytes().map_err(|e| e.to_string())?;
        
        tokio::fs::write(&output_path, &output_bytes).await.map_err(|e| e.to_string())?;
        
        return Ok(ComparisonResponse {
            stats: serde_json::json!({
                "total_changes": result.total_changes(),
            }),
            output_path,
        });
    } else if ext == "pptx" || ext == "pptm" {
        let older = PmlDocument::from_bytes(&older_bytes).map_err(|e| e.to_string())?;
        let newer = PmlDocument::from_bytes(&newer_bytes).map_err(|e| e.to_string())?;
        
        let pml_settings: PmlComparerSettings = if let Some(s) = settings {
            serde_json::from_value(s).unwrap_or_default()
        } else {
            PmlComparerSettings::default()
        };
        
        let result = PmlComparer::compare(&older, &newer, Some(&pml_settings)).map_err(|e| e.to_string())?;
        
        let marked_doc = redline_core::pml::render_marked_presentation(&newer, &result, &pml_settings).map_err(|e| e.to_string())?;
        let output_bytes = marked_doc.to_bytes().map_err(|e| e.to_string())?;
        
        tokio::fs::write(&output_path, &output_bytes).await.map_err(|e| e.to_string())?;
        
        return Ok(ComparisonResponse {
            stats: serde_json::json!({
                "total_changes": result.total_changes,
            }),
            output_path,
        });
    }

    Err("Unsupported file extension".to_string())
}


pub fn init<R: Runtime>() -> TauriPlugin<R> {
    Builder::new("redline")
        .invoke_handler(tauri::generate_handler![compare_documents])
        .build()
}
