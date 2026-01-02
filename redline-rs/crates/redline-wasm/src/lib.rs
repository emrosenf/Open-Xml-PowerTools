use serde::Serialize;
use wasm_bindgen::prelude::*;

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

    let older_doc =
        redline_core::WmlDocument::from_bytes(older).map_err(|e| JsError::new(&e.to_string()))?;
    let newer_doc =
        redline_core::WmlDocument::from_bytes(newer).map_err(|e| JsError::new(&e.to_string()))?;

    let result = redline_core::WmlComparer::compare(&older_doc, &newer_doc, settings.as_ref())
        .map_err(|e| JsError::new(&e.to_string()))?;

    let result = CompareResult {
        insertions: result.insertions,
        deletions: result.deletions,
        total_revisions: result.revision_count,
    };

    serde_wasm_bindgen::to_value(&result).map_err(|e| JsError::new(&e.to_string()))
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

    let older_doc =
        redline_core::WmlDocument::from_bytes(older).map_err(|e| JsError::new(&e.to_string()))?;
    let newer_doc =
        redline_core::WmlDocument::from_bytes(newer).map_err(|e| JsError::new(&e.to_string()))?;

    let result = redline_core::WmlComparer::compare(&older_doc, &newer_doc, settings.as_ref())
        .map_err(|e| JsError::new(&e.to_string()))?;

    Ok(result.document)
}

#[derive(Serialize)]
pub struct CompareResultWithChanges {
    #[serde(with = "serde_bytes")]
    pub document: Vec<u8>,
    pub changes: Vec<redline_core::wml::WmlChange>,
    pub insertions: usize,
    pub deletions: usize,
    pub total_revisions: usize,
}

#[wasm_bindgen]
pub fn compare_word_documents_with_changes(
    older: &[u8],
    newer: &[u8],
    settings_json: Option<String>,
) -> Result<JsValue, JsError> {
    let settings: Option<redline_core::WmlComparerSettings> = settings_json
        .map(|s| serde_json::from_str(&s))
        .transpose()
        .map_err(|e: serde_json::Error| JsError::new(&e.to_string()))?;

    let older_doc =
        redline_core::WmlDocument::from_bytes(older).map_err(|e| JsError::new(&e.to_string()))?;
    let newer_doc =
        redline_core::WmlDocument::from_bytes(newer).map_err(|e| JsError::new(&e.to_string()))?;

    let result = redline_core::WmlComparer::compare(&older_doc, &newer_doc, settings.as_ref())
        .map_err(|e| JsError::new(&e.to_string()))?;

    let result_struct = CompareResultWithChanges {
        document: result.document,
        changes: result.changes,
        insertions: result.insertions,
        deletions: result.deletions,
        total_revisions: result.revision_count,
    };

    serde_wasm_bindgen::to_value(&result_struct).map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn build_change_list(
    changes_json: JsValue,
    options_json: Option<String>,
) -> Result<JsValue, JsError> {
    let changes: Vec<redline_core::wml::WmlChange> =
        serde_wasm_bindgen::from_value(changes_json).map_err(|e| JsError::new(&e.to_string()))?;

    let options: redline_core::wml::WmlChangeListOptions = options_json
        .map(|s| serde_json::from_str(&s))
        .transpose()
        .map_err(|e: serde_json::Error| JsError::new(&e.to_string()))?
        .unwrap_or_default();

    let items = redline_core::wml::build_change_list(&changes, &options);

    serde_wasm_bindgen::to_value(&items).map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn accept_revisions_by_id(document: &[u8], revision_ids: &[i32]) -> Result<Vec<u8>, JsError> {
    let doc = redline_core::WmlDocument::from_bytes(document)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let main_doc = doc
        .main_document()
        .map_err(|e| JsError::new(&e.to_string()))?;
    let root = main_doc
        .root()
        .ok_or_else(|| JsError::new("Document has no root"))?;

    let new_xml = redline_core::wml::accept_revisions_by_id(&main_doc, root, revision_ids);

    // We need to package this XML back into a DOCX
    // For now, WmlDocument doesn't support easy "replace main document part" in-place for output
    // But we can clone the original document's package and replace the part.
    // Actually, WmlComparer does this.
    // Ideally WmlDocument should expose a way to save with modified parts.
    // The current redline-core APIs mostly return Vec<u8> from WmlComparer.

    // Let's assume we want to return the FULL docx bytes.
    // WmlDocument stores the package. We need to update the package.
    // But WmlDocument::from_bytes takes ownership of bytes (or copies).
    // The `new_xml` is just the main document part (document.xml).

    // Hack: We can use the package from `doc` if we can modify it.
    // `doc.package_mut()` allows modification.

    // But `accept_revisions_by_id` returns a NEW `XmlDocument` (arena).
    // We need to serialize this `XmlDocument` to XML string/bytes and put it into the package.

    let _xml_string = redline_core::xml::builder::serialize(&new_xml)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let mut mut_doc = redline_core::WmlDocument::from_bytes(document)
        .map_err(|e| JsError::new(&e.to_string()))?;

    mut_doc
        .package_mut()
        .put_xml_part("word/document.xml", &new_xml)
        .map_err(|e| JsError::new(&e.to_string()))?;

    mut_doc.to_bytes().map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn reject_revisions_by_id(document: &[u8], revision_ids: &[i32]) -> Result<Vec<u8>, JsError> {
    let doc = redline_core::WmlDocument::from_bytes(document)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let main_doc = doc
        .main_document()
        .map_err(|e| JsError::new(&e.to_string()))?;
    let root = main_doc
        .root()
        .ok_or_else(|| JsError::new("Document has no root"))?;

    let new_xml = redline_core::wml::reject_revisions_by_id(&main_doc, root, revision_ids);

    // Similar to accept, put back into package
    let mut mut_doc = redline_core::WmlDocument::from_bytes(document)
        .map_err(|e| JsError::new(&e.to_string()))?;

    mut_doc
        .package_mut()
        .put_xml_part("word/document.xml", &new_xml)
        .map_err(|e| JsError::new(&e.to_string()))?;

    mut_doc.to_bytes().map_err(|e| JsError::new(&e.to_string()))
}

#[derive(Serialize)]
pub struct SmlCompareResultWithChanges {
    #[serde(with = "serde_bytes")]
    pub document: Vec<u8>,
    pub changes: Vec<redline_core::sml::SmlChange>,
    pub insertions: usize,
    pub deletions: usize,
    pub revision_count: usize,
}

#[wasm_bindgen]
pub fn compare_spreadsheets(
    older: &[u8],
    newer: &[u8],
    settings_json: Option<String>,
) -> Result<JsValue, JsError> {
    let settings: Option<redline_core::SmlComparerSettings> = settings_json
        .map(|s| serde_json::from_str(&s))
        .transpose()
        .map_err(|e: serde_json::Error| JsError::new(&e.to_string()))?;

    let older_doc =
        redline_core::SmlDocument::from_bytes(older).map_err(|e| JsError::new(&e.to_string()))?;
    let newer_doc =
        redline_core::SmlDocument::from_bytes(newer).map_err(|e| JsError::new(&e.to_string()))?;

    let (marked_doc, result) =
        redline_core::SmlComparer::compare_and_render(&older_doc, &newer_doc, settings.as_ref())
            .map_err(|e| JsError::new(&e.to_string()))?;

    let document_bytes = marked_doc
        .to_bytes()
        .map_err(|e| JsError::new(&e.to_string()))?;

    let insertions = result.cells_added();
    let deletions = result.cells_deleted();
    let revision_count = result.total_changes();

    let result_struct = SmlCompareResultWithChanges {
        document: document_bytes,
        changes: result.changes,
        insertions,
        deletions,
        revision_count,
    };

    serde_wasm_bindgen::to_value(&result_struct).map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn build_sml_change_list(
    changes_json: JsValue,
    options_json: Option<String>,
) -> Result<JsValue, JsError> {
    let changes: Vec<redline_core::sml::SmlChange> =
        serde_wasm_bindgen::from_value(changes_json).map_err(|e| JsError::new(&e.to_string()))?;

    let options: redline_core::sml::SmlChangeListOptions = options_json
        .map(|s| serde_json::from_str(&s))
        .transpose()
        .map_err(|e: serde_json::Error| JsError::new(&e.to_string()))?
        .unwrap_or_default();

    let items = redline_core::sml::build_change_list(&changes, &options);

    serde_wasm_bindgen::to_value(&items).map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn apply_sml_changes(document: &[u8], changes_json: JsValue) -> Result<Vec<u8>, JsError> {
    let changes: Vec<redline_core::sml::SmlChange> =
        serde_wasm_bindgen::from_value(changes_json).map_err(|e| JsError::new(&e.to_string()))?;

    redline_core::sml::apply_sml_changes(document, &changes)
        .map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn revert_sml_changes(document: &[u8], changes_json: JsValue) -> Result<Vec<u8>, JsError> {
    let changes: Vec<redline_core::sml::SmlChange> =
        serde_wasm_bindgen::from_value(changes_json).map_err(|e| JsError::new(&e.to_string()))?;

    redline_core::sml::revert_sml_changes(document, &changes)
        .map_err(|e| JsError::new(&e.to_string()))
}

#[derive(Serialize)]
pub struct PmlCompareResultWithChanges {
    #[serde(with = "serde_bytes")]
    pub document: Vec<u8>,
    pub changes: Vec<redline_core::pml::PmlChange>,
    pub insertions: usize,
    pub deletions: usize,
    pub revision_count: usize,
}

#[wasm_bindgen]
pub fn compare_presentations(
    older: &[u8],
    newer: &[u8],
    settings_json: Option<String>,
) -> Result<JsValue, JsError> {
    let settings: Option<redline_core::PmlComparerSettings> = settings_json
        .map(|s| serde_json::from_str(&s))
        .transpose()
        .map_err(|e: serde_json::Error| JsError::new(&e.to_string()))?;

    let older_doc =
        redline_core::PmlDocument::from_bytes(older).map_err(|e| JsError::new(&e.to_string()))?;
    let newer_doc =
        redline_core::PmlDocument::from_bytes(newer).map_err(|e| JsError::new(&e.to_string()))?;

    // Use compare() for now, as we don't have produce_marked_presentation exposed yet for rendering
    // But wait, m2p.5 says "return PmlComparisonResult with changes" and "WASM bindings for PML change viewer APIs"
    // AND "render_marked_presentation" is in pml/mod.rs (re-exported).

    // We should implement compare_and_render in PmlComparer if we want consistency with SML.
    // Or just use render_marked_presentation here.

    // Let's first compute the result.
    let result = redline_core::PmlComparer::compare(&older_doc, &newer_doc, settings.as_ref())
        .map_err(|e| JsError::new(&e.to_string()))?;

    // Then render the marked presentation
    let marked_doc = redline_core::pml::render_marked_presentation(
        &newer_doc,
        &result,
        settings
            .as_ref()
            .unwrap_or(&redline_core::PmlComparerSettings::default()),
    )
    .map_err(|e| JsError::new(&e.to_string()))?;

    let document_bytes = marked_doc
        .to_bytes()
        .map_err(|e| JsError::new(&e.to_string()))?;

    let result_struct = PmlCompareResultWithChanges {
        document: document_bytes,
        changes: result.changes,
        insertions: result.slides_inserted as usize + result.shapes_inserted as usize,
        deletions: result.slides_deleted as usize + result.shapes_deleted as usize,
        revision_count: result.total_changes as usize,
    };

    serde_wasm_bindgen::to_value(&result_struct).map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn build_pml_change_list(
    changes_json: JsValue,
    options_json: Option<String>,
) -> Result<JsValue, JsError> {
    let changes: Vec<redline_core::pml::PmlChange> =
        serde_wasm_bindgen::from_value(changes_json).map_err(|e| JsError::new(&e.to_string()))?;

    let options: redline_core::pml::PmlChangeListOptions = options_json
        .map(|s| serde_json::from_str(&s))
        .transpose()
        .map_err(|e: serde_json::Error| JsError::new(&e.to_string()))?
        .unwrap_or_default();

    let items = redline_core::pml::build_change_list(&changes, &options);

    serde_wasm_bindgen::to_value(&items).map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn apply_pml_changes(document: &[u8], changes_json: JsValue) -> Result<Vec<u8>, JsError> {
    let changes: Vec<redline_core::pml::PmlChange> =
        serde_wasm_bindgen::from_value(changes_json).map_err(|e| JsError::new(&e.to_string()))?;

    redline_core::pml::apply_pml_changes(document, &changes)
        .map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn revert_pml_changes(document: &[u8], changes_json: JsValue) -> Result<Vec<u8>, JsError> {
    let changes: Vec<redline_core::pml::PmlChange> =
        serde_wasm_bindgen::from_value(changes_json).map_err(|e| JsError::new(&e.to_string()))?;

    redline_core::pml::revert_pml_changes(document, &changes)
        .map_err(|e| JsError::new(&e.to_string()))
}
