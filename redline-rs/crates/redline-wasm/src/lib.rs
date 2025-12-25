use wasm_bindgen::prelude::*;

#[wasm_bindgen(start)]
pub fn init() {
    console_error_panic_hook::set_once();
}

#[wasm_bindgen]
pub fn compare_word_documents(
    older: &[u8],
    newer: &[u8],
    settings_json: Option<String>,
) -> Result<Vec<u8>, JsError> {
    let _settings: Option<redline_core::WmlComparerSettings> = settings_json
        .map(|s| serde_json::from_str(&s))
        .transpose()
        .map_err(|e| JsError::new(&e.to_string()))?;

    let older_doc = redline_core::WmlDocument::from_bytes(older)
        .map_err(|e| JsError::new(&e.to_string()))?;
    let newer_doc = redline_core::WmlDocument::from_bytes(newer)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let _result = redline_core::WmlComparer::compare(&older_doc, &newer_doc, None)
        .map_err(|e| JsError::new(&e.to_string()))?;

    todo!("WASM bindings not yet implemented - Phase 6")
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
        .map_err(|e| JsError::new(&e.to_string()))?;

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
        .map_err(|e| JsError::new(&e.to_string()))?;

    let _older_doc = redline_core::PmlDocument::from_bytes(older)
        .map_err(|e| JsError::new(&e.to_string()))?;
    let _newer_doc = redline_core::PmlDocument::from_bytes(newer)
        .map_err(|e| JsError::new(&e.to_string()))?;

    todo!("WASM bindings not yet implemented - Phase 6")
}
