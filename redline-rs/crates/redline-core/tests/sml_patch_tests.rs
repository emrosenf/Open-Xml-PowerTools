use redline_core::sml::{apply_sml_changes, SmlChange, SmlChangeType, SmlDocument};
use std::fs;
use std::path::PathBuf;

fn get_test_file_path(filename: &str) -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.pop(); // crates
    path.pop(); // redline-rs
    path.pop(); // rust-port-phase0
    path.push("TestFiles");
    path.push(filename);
    path
}

#[test]
fn test_apply_sml_changes_value() {
    // We need a simple spreadsheet to test. SH007-One-Cell-Table.xlsx might be good.
    // Or just any spreadsheet.
    let path = get_test_file_path("SH007-One-Cell-Table.xlsx");
    if !path.exists() {
        println!("Test file not found, skipping");
        return;
    }

    let bytes = fs::read(&path).expect("Failed to read file");

    // Define a change
    let change = SmlChange {
        change_type: SmlChangeType::ValueChanged,
        sheet_name: Some("Sheet1".to_string()),
        cell_address: Some("A1".to_string()),
        new_value: Some("Patched Value".to_string()),
        ..Default::default()
    };

    let patched_bytes =
        apply_sml_changes(&bytes, &[change.clone()]).expect("Failed to apply changes");

    // Verify the change
    let doc = SmlDocument::from_bytes(&patched_bytes).expect("Failed to load patched doc");
    let _retriever = redline_core::sml::SmlDataRetriever::retrieve_sheet(&doc, "Sheet1")
        .expect("Failed to retrieve sheet");

    // Test revert
    let reverted_bytes = redline_core::sml::revert_sml_changes(&patched_bytes, &[change])
        .expect("Failed to revert changes");
    let _reverted_doc =
        SmlDocument::from_bytes(&reverted_bytes).expect("Failed to load reverted doc");
}
