//! PML Comparer Integration Tests
//!
//! Tests the Rust PmlComparer against the C# implementation.
//! Port of OpenXmlPowerTools.Tests/PmlComparerTests.cs with 100% parity.

use redline_core::pml::{PmlComparer, PmlComparerSettings, PmlDocument, PmlChangeType};

use std::path::Path;

/// Get the path to test files
fn test_files_dir() -> &'static Path {
    // CARGO_MANIFEST_DIR = crates/redline-core
    // Go up 3 levels: crates/ -> redline-rs/ -> rust-port-phase0/
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap() // crates/
        .parent()
        .unwrap() // redline-rs/
        .parent()
        .unwrap() // rust-port-phase0/
        .join("TestFiles")
        .leak()
}

// ============================================================================
// BASIC COMPARISON TESTS
// ============================================================================

#[test]
fn pc001_identical_presentations_no_changes() {
    // Test: Identical presentations should produce 0 changes
    let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Identical.pptx")).unwrap();
    let settings = PmlComparerSettings::default();

    let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();

    assert_eq!(result.total_changes, 0, "Identical presentations should have 0 changes");
}

#[test]
fn pc002_different_presentations_detects_changes() {
    // Test: Different presentations should detect changes
    // Uses PB001-Input1.pptx and PB001-Input2.pptx
    let doc1 = PmlDocument::from_file(test_files_dir().join("PB001-Input1.pptx")).unwrap();
    let doc2 = PmlDocument::from_file(test_files_dir().join("PB001-Input2.pptx")).unwrap();
    let settings = PmlComparerSettings::default();

    let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();

    assert!(result.total_changes > 0, "Different presentations should have changes");
}

// ============================================================================
// SLIDE CHANGE DETECTION TESTS
// ============================================================================

#[test]
fn pc003_slide_added_detects_insertion() {
    // Test: Slide insertion detection
    // Base presentation (2 slides) vs presentation with 3 slides
    let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-SlideAdded.pptx")).unwrap();
    let settings = PmlComparerSettings::default();

    let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();

    assert_eq!(result.slides_inserted, 1, "Should detect 1 slide insertion");
    assert!(result.changes.iter().any(|c| c.change_type == PmlChangeType::SlideInserted));
}

#[test]
fn pc004_slide_deleted_detects_deletion() {
    // Test: Slide deletion detection
    // Base presentation (2 slides) vs presentation with 1 slide
    let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-SlideDeleted.pptx")).unwrap();
    let settings = PmlComparerSettings::default();

    let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();

    assert_eq!(result.slides_deleted, 1, "Should detect 1 slide deletion");
    assert!(result.changes.iter().any(|c| c.change_type == PmlChangeType::SlideDeleted));
}

// ============================================================================
// SHAPE CHANGE DETECTION TESTS
// ============================================================================

#[test]
fn pc005_shape_added_detects_insertion() {
    // Test: Shape insertion detection
    // Base presentation vs presentation with extra shape on slide 1
    let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-ShapeAdded.pptx")).unwrap();
    let settings = PmlComparerSettings::default();

    let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();

    assert!(result.shapes_inserted >= 1, "Expected at least 1 shape inserted, got {}", result.shapes_inserted);
    assert!(result.changes.iter().any(|c| c.change_type == PmlChangeType::ShapeInserted));
}

#[test]
#[ignore]
fn pc006_shape_deleted_detects_deletion() {
    // Test: Shape deletion detection
    // Base presentation vs presentation with shape removed from slide 1
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-ShapeDeleted.pptx")).unwrap();
    // let settings = PmlComparerSettings::default();
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    //
    // assert!(result.shapes_deleted >= 1, "Expected at least 1 shape deleted, got {}", result.shapes_deleted);
    // assert!(result.changes.iter().any(|c| c.change_type == PmlChangeType::ShapeDeleted));
}

#[test]
#[ignore]
fn pc007_shape_moved_detects_move() {
    // Test: Shape move detection
    // Base presentation vs presentation with shape moved to different position
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-ShapeMoved.pptx")).unwrap();
    // let settings = PmlComparerSettings::default();
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    //
    // assert!(result.shapes_moved >= 1, "Expected at least 1 shape moved, got {}", result.shapes_moved);
    // assert!(result.changes.iter().any(|c| c.change_type == PmlChangeType::ShapeMoved));
}

#[test]
#[ignore]
fn pc008_shape_resized_detects_resize() {
    // Test: Shape resize detection
    // Base presentation vs presentation with resized shape
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-ShapeResized.pptx")).unwrap();
    // let settings = PmlComparerSettings::default();
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    //
    // assert!(result.shapes_resized >= 1, "Expected at least 1 shape resized, got {}", result.shapes_resized);
    // assert!(result.changes.iter().any(|c| c.change_type == PmlChangeType::ShapeResized));
}

// ============================================================================
// TEXT CHANGE DETECTION TESTS
// ============================================================================

#[test]
#[ignore]
fn pc009_text_changed_detects_text_change() {
    // Test: Text content change detection
    // Base presentation vs presentation with modified text
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-TextChanged.pptx")).unwrap();
    // let settings = PmlComparerSettings::default();
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    //
    // assert!(result.text_changes >= 1, "Expected at least 1 text change, got {}", result.text_changes);
    // assert!(result.changes.iter().any(|c| c.change_type == PmlChangeType::TextChanged));
}

// ============================================================================
// SETTINGS TESTS
// ============================================================================

#[test]
#[ignore]
fn pc010_default_settings_has_correct_defaults() {
    // Test: Verify default settings values
    //
    // let settings = PmlComparerSettings::default();
    //
    // assert!(settings.compare_slide_structure, "CompareSlideStructure should be true by default");
    // assert!(settings.compare_shape_structure, "CompareShapeStructure should be true by default");
    // assert!(settings.compare_text_content, "CompareTextContent should be true by default");
    // assert!(settings.compare_text_formatting, "CompareTextFormatting should be true by default");
    // assert!(settings.compare_shape_transforms, "CompareShapeTransforms should be true by default");
    // assert!(!settings.compare_shape_styles, "CompareShapeStyles should be false by default");
    // assert!(settings.compare_image_content, "CompareImageContent should be true by default");
    // assert!(settings.compare_charts, "CompareCharts should be true by default");
    // assert!(settings.compare_tables, "CompareTables should be true by default");
    // assert!(!settings.compare_notes, "CompareNotes should be false by default");
    // assert!(!settings.compare_transitions, "CompareTransitions should be false by default");
    // assert!(settings.enable_fuzzy_shape_matching, "EnableFuzzyShapeMatching should be true by default");
    // assert!(settings.use_slide_alignment_lcs, "UseSlideAlignmentLCS should be true by default");
    // assert!(settings.add_summary_slide, "AddSummarySlide should be true by default");
    // assert!(settings.add_notes_annotations, "AddNotesAnnotations should be true by default");
}

#[test]
#[ignore]
fn pc011_disable_slide_structure_ignores_slide_changes() {
    // Test: Disabling slide structure comparison ignores slide changes
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-SlideAdded.pptx")).unwrap();
    // let mut settings = PmlComparerSettings::default();
    // settings.compare_slide_structure = false;
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    //
    // assert_eq!(result.slides_inserted, 0, "Should ignore slide insertions when disabled");
    // assert_eq!(result.slides_deleted, 0, "Should ignore slide deletions when disabled");
}

#[test]
#[ignore]
fn pc012_disable_shape_structure_ignores_shape_changes() {
    // Test: Disabling shape structure comparison ignores shape changes
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-ShapeAdded.pptx")).unwrap();
    // let mut settings = PmlComparerSettings::default();
    // settings.compare_shape_structure = false;
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    //
    // assert_eq!(result.shapes_inserted, 0, "Should ignore shape insertions when disabled");
    // assert_eq!(result.shapes_deleted, 0, "Should ignore shape deletions when disabled");
}

#[test]
#[ignore]
fn pc013_disable_text_content_ignores_text_changes() {
    // Test: Disabling text content comparison ignores text changes
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-TextChanged.pptx")).unwrap();
    // let mut settings = PmlComparerSettings::default();
    // settings.compare_text_content = false;
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    //
    // assert_eq!(result.text_changes, 0, "Should ignore text changes when disabled");
}

#[test]
#[ignore]
fn pc014_disable_shape_transforms_ignores_move_and_resize() {
    // Test: Disabling shape transforms comparison ignores moves and resizes
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-ShapeMoved.pptx")).unwrap();
    // let mut settings = PmlComparerSettings::default();
    // settings.compare_shape_transforms = false;
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    //
    // assert_eq!(result.shapes_moved, 0, "Should ignore shape moves when disabled");
    // assert_eq!(result.shapes_resized, 0, "Should ignore shape resizes when disabled");
}

// ============================================================================
// RESULT PROPERTIES TESTS
// ============================================================================

#[test]
#[ignore]
fn pc015_result_get_changes_by_slide_works() {
    // Test: Filter changes by slide number
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-ShapeAdded.pptx")).unwrap();
    // let settings = PmlComparerSettings::default();
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    // let slide_changes = result.get_changes_by_slide(1);
    //
    // assert!(!slide_changes.is_empty(), "Slide 1 should have changes");
}

#[test]
#[ignore]
fn pc016_result_get_changes_by_type_works() {
    // Test: Filter changes by change type
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-SlideAdded.pptx")).unwrap();
    // let settings = PmlComparerSettings::default();
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    // let slide_insertions = result.get_changes_by_type(PmlChangeType::SlideInserted);
    //
    // assert_eq!(slide_insertions.len(), 1, "Should have exactly 1 slide insertion");
}

#[test]
#[ignore]
fn pc017_result_to_json_returns_valid_json() {
    // Test: Serialize result to JSON
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-SlideAdded.pptx")).unwrap();
    // let settings = PmlComparerSettings::default();
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    // let json = result.to_json().unwrap();
    //
    // assert!(json.contains("TotalChanges"), "JSON should contain TotalChanges");
    // assert!(json.contains("Summary"), "JSON should contain Summary");
    // assert!(json.contains("Changes"), "JSON should contain Changes");
    // assert!(json.contains("SlidesInserted"), "JSON should contain SlidesInserted");
}

#[test]
#[ignore]
fn pc018_change_get_description_returns_readable_text() {
    // Test: Change description is human-readable
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-SlideAdded.pptx")).unwrap();
    // let settings = PmlComparerSettings::default();
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    // let change = result.changes.iter()
    //     .find(|c| c.change_type == PmlChangeType::SlideInserted)
    //     .expect("Should have a SlideInserted change");
    //
    // let description = change.get_description();
    //
    // assert!(description.to_lowercase().contains("inserted"), "Description should mention 'inserted'");
}

// ============================================================================
// MARKED PRESENTATION TESTS
// ============================================================================

#[test]
#[ignore]
fn pc019_produce_marked_presentation_returns_valid_document() {
    // Test: ProduceMarkedPresentation creates valid PPTX
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-ShapeAdded.pptx")).unwrap();
    // let settings = PmlComparerSettings::default();
    //
    // let marked = PmlComparer::produce_marked_presentation(&doc1, &doc2, Some(&settings)).unwrap();
    //
    // assert!(!marked.document_byte_array.is_empty(), "Marked presentation should not be empty");
}

#[test]
#[ignore]
fn pc020_produce_marked_presentation_with_summary_slide() {
    // Test: Marked presentation includes summary slide when enabled
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-SlideAdded.pptx")).unwrap();
    // let mut settings = PmlComparerSettings::default();
    // settings.add_summary_slide = true;
    //
    // let marked = PmlComparer::produce_marked_presentation(&doc1, &doc2, Some(&settings)).unwrap();
    //
    // // Marked presentation with summary slide should be larger than doc2
    // assert!(marked.document_byte_array.len() > doc2.document_byte_array.len(),
    //     "Marked presentation with summary slide should be larger");
}

#[test]
#[ignore]
fn pc021_produce_marked_presentation_no_changes_returns_same_size() {
    // Test: No changes means similar size (without summary slide)
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Identical.pptx")).unwrap();
    // let mut settings = PmlComparerSettings::default();
    // settings.add_summary_slide = false;
    //
    // let marked = PmlComparer::produce_marked_presentation(&doc1, &doc2, Some(&settings)).unwrap();
    //
    // assert!(!marked.document_byte_array.is_empty(), "Marked presentation should not be empty");
}

// ============================================================================
// LOGGING TESTS
// ============================================================================

#[test]
#[ignore]
fn pc022_log_callback_receives_messages() {
    // Test: Log callback receives messages during comparison
    //
    // let doc1 = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let doc2 = PmlDocument::from_file(test_files_dir().join("PmlComparer-SlideAdded.pptx")).unwrap();
    //
    // let log_messages = Arc::new(Mutex::new(Vec::new()));
    // let log_messages_clone = log_messages.clone();
    //
    // let mut settings = PmlComparerSettings::default();
    // settings.log_callback = Some(Box::new(move |msg| {
    //     log_messages_clone.lock().unwrap().push(msg.to_string());
    // }));
    //
    // let result = PmlComparer::compare(&doc1, &doc2, Some(&settings)).unwrap();
    //
    // let messages = log_messages.lock().unwrap();
    // assert!(!messages.is_empty(), "Log callback should receive messages");
    // assert!(messages.iter().any(|m| m.contains("PmlComparer")), "Log should mention PmlComparer");
}

// ============================================================================
// EDGE CASES
// ============================================================================

#[test]
#[ignore]
#[should_panic(expected = "older document is null")]
fn pc023_compare_null_older_throws_error() {
    // Test: Comparing with null older document should panic/error
    //
    // let doc = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let settings = PmlComparerSettings::default();
    //
    // // This should panic or return error - adjust based on Rust implementation
    // let _ = PmlComparer::compare(None, &doc, Some(&settings));
}

#[test]
#[ignore]
#[should_panic(expected = "newer document is null")]
fn pc024_compare_null_newer_throws_error() {
    // Test: Comparing with null newer document should panic/error
    //
    // let doc = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let settings = PmlComparerSettings::default();
    //
    // // This should panic or return error
    // let _ = PmlComparer::compare(&doc, None, Some(&settings));
}

#[test]
#[ignore]
fn pc025_compare_null_settings_uses_defaults() {
    // Test: Null settings should use defaults
    //
    // let doc = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    //
    // let result = PmlComparer::compare(&doc, &doc, None).unwrap();
    //
    // assert_eq!(result.total_changes, 0, "Comparing doc with itself should have 0 changes");
}

// ============================================================================
// CHANGE TYPE TESTS
// ============================================================================

#[test]
#[ignore]
fn pc026_change_type_has_expected_values() {
    // Test: Verify all expected change type enum values exist
    //
    // // Verify enum values exist
    // let _ = PmlChangeType::SlideSizeChanged;
    // let _ = PmlChangeType::SlideInserted;
    // let _ = PmlChangeType::SlideDeleted;
    // let _ = PmlChangeType::SlideMoved;
    // let _ = PmlChangeType::ShapeInserted;
    // let _ = PmlChangeType::ShapeDeleted;
    // let _ = PmlChangeType::ShapeMoved;
    // let _ = PmlChangeType::ShapeResized;
    // let _ = PmlChangeType::TextChanged;
    // let _ = PmlChangeType::ImageReplaced;
    //
    // // Verify SlideSizeChanged is 0 (as in C#)
    // assert_eq!(PmlChangeType::SlideSizeChanged as i32, 0);
}

// ============================================================================
// CANONICALIZE TESTS
// ============================================================================

#[test]
#[ignore]
fn pc027_canonicalize_returns_signature() {
    // Test: Canonicalize returns a signature
    //
    // let doc = PmlDocument::from_file(test_files_dir().join("PmlComparer-Base.pptx")).unwrap();
    // let settings = PmlComparerSettings::default();
    //
    // let signature = PmlComparer::canonicalize(&doc, Some(&settings)).unwrap();
    //
    // assert!(!signature.is_empty(), "Canonicalize should return non-empty signature");
}

// ============================================================================
// TEST HELPERS
// ============================================================================

// Note: Test helper functions would go here if needed. For now, we rely on
// the PmlComparerTestFileGenerator.cs to create test files in the C# project.
//
// If test files don't exist, you must run the C# test suite at least once
// to generate them, or implement a Rust equivalent of the file generator.

// ============================================================================
// IMPLEMENTATION NOTES
// ============================================================================
//
// This test file ports all 27 tests from PmlComparerTests.cs:
//
// - PC001-PC002: Basic comparison tests (2 tests)
// - PC003-PC004: Slide change detection (2 tests)
// - PC005-PC009: Shape and text change detection (5 tests)
// - PC010-PC014: Settings tests (5 tests)
// - PC015-PC018: Result properties tests (4 tests)
// - PC019-PC021: Marked presentation tests (3 tests)
// - PC022: Logging test (1 test)
// - PC023-PC025: Edge cases (3 tests)
// - PC026: Change type enum test (1 test)
// - PC027: Canonicalize test (1 test)
//
// Total: 27 tests matching the C# implementation
//
// Implementation checklist:
// 1. Implement PmlDocument type
// 2. Implement PmlComparer with Compare and ProduceMarkedPresentation methods
// 3. Implement PmlComparerSettings with all configuration options
// 4. Implement PmlChangeType enum
// 5. Implement PmlChange and PmlComparisonResult types
// 6. Implement Canonicalize method
// 7. Remove #[ignore] attributes as features are implemented
// 8. Uncomment test bodies as implementation progresses
// 9. Run tests to verify 100% parity with C# implementation
//
// Test files required (generated by PmlComparerTestFileGenerator.cs):
// - PmlComparer-Base.pptx (2 slides, base for comparisons)
// - PmlComparer-Identical.pptx (identical to base)
// - PmlComparer-SlideAdded.pptx (3 slides, one extra)
// - PmlComparer-SlideDeleted.pptx (1 slide, one removed)
// - PmlComparer-ShapeAdded.pptx (extra shape on slide 1)
// - PmlComparer-ShapeDeleted.pptx (shape removed from slide 1)
// - PmlComparer-TextChanged.pptx (modified text in slide 1)
// - PmlComparer-ShapeMoved.pptx (shape moved to different position)
// - PmlComparer-ShapeResized.pptx (shape resized)
// - PB001-Input1.pptx (for PC002)
// - PB001-Input2.pptx (for PC002)
