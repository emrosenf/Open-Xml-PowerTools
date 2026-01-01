use redline_core::wml::{accept_revisions_by_id, reject_revisions_by_id, WmlChangeType};
use redline_core::{WmlComparer, WmlComparerSettings, WmlDocument};
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
fn test_populate_wml_changes() {
    // Use simple documents that should have changes
    let left_path = get_test_file_path("WC/WC002-Unmodified.docx");
    let right_path = get_test_file_path("WC/WC002-InsertInMiddle.docx");

    if !left_path.exists() || !right_path.exists() {
        println!("Test files not found, skipping test_populate_wml_changes");
        return;
    }

    let left_bytes = fs::read(&left_path).expect("Failed to read left file");
    let right_bytes = fs::read(&right_path).expect("Failed to read right file");

    let doc1 = WmlDocument::from_bytes(&left_bytes).expect("Failed to load left doc");
    let doc2 = WmlDocument::from_bytes(&right_bytes).expect("Failed to load right doc");

    let settings = WmlComparerSettings::default();
    let result = WmlComparer::compare(&doc1, &doc2, Some(&settings)).expect("Compare failed");

    // Verify changes were populated
    // We expect some changes for different documents
    assert!(
        !result.changes.is_empty(),
        "Changes vector should not be empty"
    );

    // Check first change structure
    let first = &result.changes[0];
    println!("First change: {:?}", first);

    assert!(first.revision_id > 0, "Revision ID should be positive");
    match first.change_type {
        WmlChangeType::TextInserted | WmlChangeType::TextDeleted | WmlChangeType::FormatChanged => {
            // Good
        }
        _ => panic!("Unexpected change type: {:?}", first.change_type),
    }

    if first.change_type == WmlChangeType::TextInserted {
        assert!(
            first.new_text.is_some(),
            "TextInserted should have new_text"
        );
    } else if first.change_type == WmlChangeType::TextDeleted {
        assert!(first.old_text.is_some(), "TextDeleted should have old_text");
    }
}

#[test]
fn test_accept_reject_by_id() {
    let path = get_test_file_path("RA001-Tracked-Revisions-01.docx");
    if !path.exists() {
        println!("Test file RA001-Tracked-Revisions-01.docx not found, skipping");
        return;
    }

    let bytes = fs::read(&path).expect("Failed to read file");
    let doc = WmlDocument::from_bytes(&bytes).expect("Failed to load doc");
    let main_doc = doc.main_document().expect("No main document");
    let root = main_doc.root().expect("No root");

    // Test 1: Accept non-existent ID
    let doc_partial = accept_revisions_by_id(&main_doc, root, &[999999]);
    // The document should still be valid
    assert!(doc_partial.root().is_some());

    // Test 2: Reject non-existent ID
    let doc_rejected = reject_revisions_by_id(&main_doc, root, &[999999]);
    assert!(doc_rejected.root().is_some());
}
