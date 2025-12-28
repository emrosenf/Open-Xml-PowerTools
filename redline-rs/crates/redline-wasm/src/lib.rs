use wasm_bindgen::prelude::*;
use serde::Serialize;

#[wasm_bindgen(start)]
pub fn init() {
    console_error_panic_hook::set_once();
}

/// Result of comparing two Word documents
#[derive(Serialize)]
pub struct CompareResult {
    /// Number of insertions detected
    pub insertions: usize,
    /// Number of deletions detected  
    pub deletions: usize,
    /// Total revision count
    pub total_revisions: usize,
}

/// Compare two Word documents and return revision statistics (without generating output document)
#[wasm_bindgen]
pub fn count_word_revisions(
    older: &[u8],
    newer: &[u8],
    settings_json: Option<String>,
) -> Result<JsValue, JsError> {
    let settings: Option<redline_core::WmlComparerSettings> = settings_json
        .map(|s| serde_json::from_str(&s))
        .transpose()
        .map_err(|e: serde_json::Error| JsError::new(&e.to_string()))?;

    let older_doc = redline_core::WmlDocument::from_bytes(older)
        .map_err(|e| JsError::new(&e.to_string()))?;
    let newer_doc = redline_core::WmlDocument::from_bytes(newer)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let result = redline_core::WmlComparer::compare(&older_doc, &newer_doc, settings.as_ref())
        .map_err(|e| JsError::new(&e.to_string()))?;

    let result = CompareResult {
        insertions: result.insertions,
        deletions: result.deletions,
        total_revisions: result.revision_count,
    };

    serde_wasm_bindgen::to_value(&result)
        .map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn compare_word_documents(
    older: &[u8],
    newer: &[u8],
    settings_json: Option<String>,
) -> Result<Vec<u8>, JsError> {
    let settings: Option<redline_core::WmlComparerSettings> = settings_json
        .map(|s| serde_json::from_str(&s))
        .transpose()
        .map_err(|e: serde_json::Error| JsError::new(&e.to_string()))?;

    let older_doc = redline_core::WmlDocument::from_bytes(older)
        .map_err(|e| JsError::new(&e.to_string()))?;
    let newer_doc = redline_core::WmlDocument::from_bytes(newer)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let result = redline_core::WmlComparer::compare(&older_doc, &newer_doc, settings.as_ref())
        .map_err(|e| JsError::new(&e.to_string()))?;

    // Return the compared document as bytes
    Ok(result.document)
}

#[wasm_bindgen]
pub fn compare_spreadsheets(
    older: &[u8],
    newer: &[u8],
    settings_json: Option<String>,
) -> Result<JsValue, JsError> {
    let _settings: Option<redline_core::SmlComparerSettings> = settings_json
        .map(|s| serde_json::from_str(&s))
        .transpose()
        .map_err(|e: serde_json::Error| JsError::new(&e.to_string()))?;

    let _older_doc = redline_core::SmlDocument::from_bytes(older)
        .map_err(|e| JsError::new(&e.to_string()))?;
    let _newer_doc = redline_core::SmlDocument::from_bytes(newer)
        .map_err(|e| JsError::new(&e.to_string()))?;

    todo!("WASM bindings not yet implemented - Phase 6")
}

#[wasm_bindgen]
pub fn compare_presentations(
    older: &[u8],
    newer: &[u8],
    settings_json: Option<String>,
) -> Result<JsValue, JsError> {
    let _settings: Option<redline_core::PmlComparerSettings> = settings_json
        .map(|s| serde_json::from_str(&s))
        .transpose()
        .map_err(|e: serde_json::Error| JsError::new(&e.to_string()))?;

    let _older_doc = redline_core::PmlDocument::from_bytes(older)
        .map_err(|e| JsError::new(&e.to_string()))?;
    let _newer_doc = redline_core::PmlDocument::from_bytes(newer)
        .map_err(|e| JsError::new(&e.to_string()))?;

    todo!("WASM bindings not yet implemented - Phase 6")
}
