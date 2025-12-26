//! WML Comparer Integration Tests
//!
//! Tests the Rust WmlComparer against expected revision counts from the C# golden files.
//! All 104 test cases from WmlComparerTests.cs are included.

use redline_core::wml::{WmlComparer, WmlComparerSettings, WmlDocument};
use std::fs;
use std::path::Path;

/// Get the path to test files
fn test_files_dir() -> &'static Path {
    // CARGO_MANIFEST_DIR = crates/redline-core
    // Go up 3 levels: crates/ -> redline-rs/ -> rust-port-phase0/
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent().unwrap()  // crates/
        .parent().unwrap()  // redline-rs/
        .parent().unwrap()  // rust-port-phase0/
        .join("TestFiles")
        .leak()
}

/// Load a test document
fn load_test_doc(relative_path: &str) -> WmlDocument {
    let path = test_files_dir().join(relative_path);
    let bytes = fs::read(&path).unwrap_or_else(|e| panic!("Failed to read {}: {}", path.display(), e));
    WmlDocument::from_bytes(&bytes).unwrap_or_else(|e| panic!("Failed to parse {}: {}", path.display(), e))
}

/// Run a comparison test
fn run_comparison_test(test_id: &str, source1: &str, source2: &str, expected_revisions: usize) {
    let doc1 = load_test_doc(source1);
    let doc2 = load_test_doc(source2);
    let settings = WmlComparerSettings::default();
    
    let result = WmlComparer::compare(&doc1, &doc2, Some(&settings))
        .unwrap_or_else(|e| panic!("{}: Comparison failed: {}", test_id, e));
    
    let actual = result.revision_count;
    
    println!("{}: Expected {}, got {} ({} ins, {} del)", 
        test_id, expected_revisions, actual, result.insertions, result.deletions);
    
    assert_eq!(
        actual, expected_revisions,
        "{}: Expected {} revisions, got {} ({} insertions, {} deletions)",
        test_id, expected_revisions, actual, result.insertions, result.deletions
    );
}

// ============================================================================
// Basic Text Comparisons
// ============================================================================

#[test]
fn wc_1000_plain() {
    run_comparison_test("WC-1000", "CA/CA001-Plain.docx", "CA/CA001-Plain-Mod.docx", 1);
}

#[test]
fn wc_1010_digits() {
    run_comparison_test("WC-1010", "WC/WC001-Digits.docx", "WC/WC001-Digits-Mod.docx", 4);
}

#[test]
fn wc_1020_deleted_paragraph() {
    run_comparison_test("WC-1020", "WC/WC001-Digits.docx", "WC/WC001-Digits-Deleted-Paragraph.docx", 1);
}

#[test]
fn wc_1030_inserted_paragraph() {
    run_comparison_test("WC-1030", "WC/WC001-Digits-Deleted-Paragraph.docx", "WC/WC001-Digits.docx", 1);
}

#[test]
fn wc_1040_diff_in_middle() {
    run_comparison_test("WC-1040", "WC/WC002-Unmodified.docx", "WC/WC002-DiffInMiddle.docx", 2);
}

#[test]
fn wc_1050_diff_at_beginning() {
    run_comparison_test("WC-1050", "WC/WC002-Unmodified.docx", "WC/WC002-DiffAtBeginning.docx", 2);
}

#[test]
fn wc_1060_delete_at_beginning() {
    run_comparison_test("WC-1060", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtBeginning.docx", 1);
}

#[test]
fn wc_1070_insert_at_beginning() {
    run_comparison_test("WC-1070", "WC/WC002-Unmodified.docx", "WC/WC002-InsertAtBeginning.docx", 1);
}

#[test]
fn wc_1080_insert_at_end() {
    run_comparison_test("WC-1080", "WC/WC002-Unmodified.docx", "WC/WC002-InsertAtEnd.docx", 1);
}

#[test]
fn wc_1090_delete_at_end() {
    run_comparison_test("WC-1090", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtEnd.docx", 1);
}

#[test]
fn wc_1100_delete_in_middle() {
    run_comparison_test("WC-1100", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteInMiddle.docx", 1);
}

#[test]
fn wc_1110_insert_in_middle() {
    run_comparison_test("WC-1110", "WC/WC002-Unmodified.docx", "WC/WC002-InsertInMiddle.docx", 1);
}

#[test]
fn wc_1120_reverse_delete() {
    run_comparison_test("WC-1120", "WC/WC002-DeleteInMiddle.docx", "WC/WC002-Unmodified.docx", 1);
}

// ============================================================================
// Table Tests
// ============================================================================

#[test]
fn wc_1140_table_delete_row() {
    run_comparison_test("WC-1140", "WC/WC006-Table.docx", "WC/WC006-Table-Delete-Row.docx", 1);
}

#[test]
fn wc_1150_table_insert_row() {
    run_comparison_test("WC-1150", "WC/WC006-Table-Delete-Row.docx", "WC/WC006-Table.docx", 1);
}

#[test]
fn wc_1160_table_delete_contents_of_row() {
    run_comparison_test("WC-1160", "WC/WC006-Table.docx", "WC/WC006-Table-Delete-Contests-of-Row.docx", 2);
}

#[test]
fn wc_1170_longest_at_end() {
    run_comparison_test("WC-1170", "WC/WC007-Unmodified.docx", "WC/WC007-Longest-At-End.docx", 2);
}

#[test]
fn wc_1180_deleted_at_beginning_of_para() {
    run_comparison_test("WC-1180", "WC/WC007-Unmodified.docx", "WC/WC007-Deleted-at-Beginning-of-Para.docx", 1);
}

#[test]
fn wc_1190_moved_into_table() {
    run_comparison_test("WC-1190", "WC/WC007-Unmodified.docx", "WC/WC007-Moved-into-Table.docx", 2);
}

#[test]
fn wc_1200_table_cell_mod() {
    run_comparison_test("WC-1200", "WC/WC009-Table-Unmodified.docx", "WC/WC009-Table-Cell-1-1-Mod.docx", 1);
}

#[test]
fn wc_1210_para_before_table() {
    run_comparison_test("WC-1210", "WC/WC010-Para-Before-Table-Unmodified.docx", "WC/WC010-Para-Before-Table-Mod.docx", 3);
}

#[test]
fn wc_1220_before_after() {
    run_comparison_test("WC-1220", "WC/WC011-Before.docx", "WC/WC011-After.docx", 2);
}

// ============================================================================
// Math Content
// ============================================================================

#[test]
fn wc_1230_math() {
    run_comparison_test("WC-1230", "WC/WC012-Math-Before.docx", "WC/WC012-Math-After.docx", 2);
}

// ============================================================================
// Images
// ============================================================================

#[test]
fn wc_1240_image() {
    run_comparison_test("WC-1240", "WC/WC013-Image-Before.docx", "WC/WC013-Image-After.docx", 2);
}

#[test]
fn wc_1250_image2() {
    run_comparison_test("WC-1250", "WC/WC013-Image-Before.docx", "WC/WC013-Image-After2.docx", 2);
}

#[test]
fn wc_1260_image3() {
    run_comparison_test("WC-1260", "WC/WC013-Image-Before2.docx", "WC/WC013-Image-After2.docx", 2);
}

// ============================================================================
// SmartArt
// ============================================================================

#[test]
fn wc_1270_smartart() {
    run_comparison_test("WC-1270", "WC/WC014-SmartArt-Before.docx", "WC/WC014-SmartArt-After.docx", 2);
}

#[test]
fn wc_1280_smartart_with_image() {
    run_comparison_test("WC-1280", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-After.docx", 2);
}

#[test]
fn wc_1310_smartart_deleted() {
    run_comparison_test("WC-1310", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After.docx", 3);
}

#[test]
fn wc_1320_smartart_deleted2() {
    run_comparison_test("WC-1320", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After2.docx", 1);
}

// ============================================================================
// Multi-paragraph
// ============================================================================

#[test]
fn wc_1330_three_paragraphs() {
    run_comparison_test("WC-1330", "WC/WC015-Three-Paragraphs.docx", "WC/WC015-Three-Paragraphs-After.docx", 3);
}

// ============================================================================
// Images with paragraphs
// ============================================================================

#[test]
fn wc_1340_para_image_para() {
    run_comparison_test("WC-1340", "WC/WC016-Para-Image-Para.docx", "WC/WC016-Para-Image-Para-w-Deleted-Image.docx", 1);
}

#[test]
fn wc_1350_image_after() {
    run_comparison_test("WC-1350", "WC/WC017-Image.docx", "WC/WC017-Image-After.docx", 3);
}

// ============================================================================
// Fields
// ============================================================================

#[test]
fn wc_1360_field_simple() {
    run_comparison_test("WC-1360", "WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-1.docx", 2);
}

#[test]
fn wc_1370_field_simple2() {
    run_comparison_test("WC-1370", "WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-2.docx", 3);
}

// ============================================================================
// Hyperlinks
// ============================================================================

#[test]
fn wc_1380_hyperlink() {
    run_comparison_test("WC-1380", "WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-1.docx", 3);
}

#[test]
fn wc_1390_hyperlink2() {
    run_comparison_test("WC-1390", "WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-2.docx", 5);
}

// ============================================================================
// Footnotes
// ============================================================================

#[test]
fn wc_1400_footnote() {
    run_comparison_test("WC-1400", "WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-1.docx", 3);
}

#[test]
fn wc_1410_footnote2() {
    run_comparison_test("WC-1410", "WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-2.docx", 5);
}

// ============================================================================
// Complex math
// ============================================================================

#[test]
fn wc_1420_math_complex() {
    run_comparison_test("WC-1420", "WC/WC021-Math-Before-1.docx", "WC/WC021-Math-After-1.docx", 9);
}

#[test]
fn wc_1430_math_complex2() {
    run_comparison_test("WC-1430", "WC/WC021-Math-Before-2.docx", "WC/WC021-Math-After-2.docx", 6);
}

#[test]
fn wc_1440_image_math_para() {
    run_comparison_test("WC-1440", "WC/WC022-Image-Math-Para-Before.docx", "WC/WC022-Image-Math-Para-After.docx", 10);
}

// ============================================================================
// Tables with images
// ============================================================================

#[test]
fn wc_1450_table_4_row_image() {
    run_comparison_test("WC-1450", "WC/WC023-Table-4-Row-Image-Before.docx", "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx", 7);
}

#[test]
fn wc_1460_table() {
    run_comparison_test("WC-1460", "WC/WC024-Table-Before.docx", "WC/WC024-Table-After.docx", 1);
}

#[test]
fn wc_1470_table2() {
    run_comparison_test("WC-1470", "WC/WC024-Table-Before.docx", "WC/WC024-Table-After2.docx", 7);
}

#[test]
fn wc_1480_simple_table() {
    run_comparison_test("WC-1480", "WC/WC025-Simple-Table-Before.docx", "WC/WC025-Simple-Table-After.docx", 4);
}

#[test]
fn wc_1500_long_table() {
    run_comparison_test("WC-1500", "WC/WC026-Long-Table-Before.docx", "WC/WC026-Long-Table-After-1.docx", 2);
}

// ============================================================================
// Twenty paragraphs
// ============================================================================

#[test]
fn wc_1510_twenty_paras() {
    run_comparison_test("WC-1510", "WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-1.docx", 2);
}

#[test]
fn wc_1520_twenty_paras_reverse() {
    run_comparison_test("WC-1520", "WC/WC027-Twenty-Paras-After-1.docx", "WC/WC027-Twenty-Paras-Before.docx", 2);
}

#[test]
fn wc_1530_twenty_paras2() {
    run_comparison_test("WC-1530", "WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-2.docx", 4);
}

// ============================================================================
// Image and math combinations
// ============================================================================

#[test]
fn wc_1540_image_math() {
    run_comparison_test("WC-1540", "WC/WC030-Image-Math-Before.docx", "WC/WC030-Image-Math-After.docx", 2);
}

#[test]
fn wc_1550_two_maths() {
    run_comparison_test("WC-1550", "WC/WC031-Two-Maths-Before.docx", "WC/WC031-Two-Maths-After.docx", 4);
}

// ============================================================================
// Paragraph properties
// ============================================================================

#[test]
fn wc_1560_para_props() {
    run_comparison_test("WC-1560", "WC/WC032-Para-with-Para-Props.docx", "WC/WC032-Para-with-Para-Props-After.docx", 3);
}

// ============================================================================
// Merged cells
// ============================================================================

#[test]
fn wc_1570_merged_cells() {
    run_comparison_test("WC-1570", "WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After1.docx", 2);
}

#[test]
fn wc_1580_merged_cells2() {
    run_comparison_test("WC-1580", "WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After2.docx", 4);
}

// ============================================================================
// Footnotes variants
// ============================================================================

#[test]
fn wc_1600_footnotes() {
    run_comparison_test("WC-1600", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After1.docx", 1);
}

#[test]
fn wc_1610_footnotes2() {
    run_comparison_test("WC-1610", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After2.docx", 4);
}

#[test]
fn wc_1620_footnotes3() {
    run_comparison_test("WC-1620", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After3.docx", 3);
}

#[test]
fn wc_1630_footnotes_reverse() {
    run_comparison_test("WC-1630", "WC/WC034-Footnotes-After3.docx", "WC/WC034-Footnotes-Before.docx", 3);
}

#[test]
fn wc_1640_footnote() {
    run_comparison_test("WC-1640", "WC/WC035-Footnote-Before.docx", "WC/WC035-Footnote-After.docx", 2);
}

#[test]
fn wc_1650_footnote_reverse() {
    run_comparison_test("WC-1650", "WC/WC035-Footnote-After.docx", "WC/WC035-Footnote-Before.docx", 2);
}

#[test]
fn wc_1660_footnote_with_table() {
    run_comparison_test("WC-1660", "WC/WC036-Footnote-With-Table-Before.docx", "WC/WC036-Footnote-With-Table-After.docx", 5);
}

#[test]
fn wc_1670_footnote_with_table_reverse() {
    run_comparison_test("WC-1670", "WC/WC036-Footnote-With-Table-After.docx", "WC/WC036-Footnote-With-Table-Before.docx", 5);
}

// ============================================================================
// Endnotes
// ============================================================================

#[test]
fn wc_1680_endnotes() {
    run_comparison_test("WC-1680", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After1.docx", 1);
}

#[test]
fn wc_1700_endnotes2() {
    run_comparison_test("WC-1700", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After2.docx", 4);
}

#[test]
fn wc_1710_endnotes3() {
    run_comparison_test("WC-1710", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After3.docx", 7);
}

#[test]
fn wc_1720_endnotes_reverse() {
    run_comparison_test("WC-1720", "WC/WC034-Endnotes-After3.docx", "WC/WC034-Endnotes-Before.docx", 7);
}

#[test]
fn wc_1730_endnote() {
    run_comparison_test("WC-1730", "WC/WC035-Endnote-Before.docx", "WC/WC035-Endnote-After.docx", 2);
}

#[test]
fn wc_1740_endnote_reverse() {
    run_comparison_test("WC-1740", "WC/WC035-Endnote-After.docx", "WC/WC035-Endnote-Before.docx", 2);
}

#[test]
fn wc_1750_endnote_with_table() {
    run_comparison_test("WC-1750", "WC/WC036-Endnote-With-Table-Before.docx", "WC/WC036-Endnote-With-Table-After.docx", 6);
}

#[test]
fn wc_1760_endnote_with_table_reverse() {
    run_comparison_test("WC-1760", "WC/WC036-Endnote-With-Table-After.docx", "WC/WC036-Endnote-With-Table-Before.docx", 6);
}

// ============================================================================
// Textboxes
// ============================================================================

#[test]
fn wc_1770_textbox() {
    run_comparison_test("WC-1770", "WC/WC037-Textbox-Before.docx", "WC/WC037-Textbox-After1.docx", 2);
}

// ============================================================================
// Line breaks
// ============================================================================

#[test]
fn wc_1780_line_breaks() {
    run_comparison_test("WC-1780", "WC/WC038-Document-With-BR-Before.docx", "WC/WC038-Document-With-BR-After.docx", 2);
}

// ============================================================================
// Revision consolidation
// ============================================================================

#[test]
fn wc_1800_revision_consolidation() {
    run_comparison_test("WC-1800", "RC/RC001-Before.docx", "RC/RC001-After1.docx", 2);
}

#[test]
fn wc_1810_revision_consolidation_image() {
    run_comparison_test("WC-1810", "RC/RC002-Image.docx", "RC/RC002-Image-After1.docx", 1);
}

// ============================================================================
// Breaks in rows
// ============================================================================

#[test]
fn wc_1820_break_in_row() {
    run_comparison_test("WC-1820", "WC/WC039-Break-In-Row.docx", "WC/WC039-Break-In-Row-After1.docx", 1);
}

// ============================================================================
// More tables
// ============================================================================

#[test]
fn wc_1830_table_5() {
    run_comparison_test("WC-1830", "WC/WC041-Table-5.docx", "WC/WC041-Table-5-Mod.docx", 2);
}

#[test]
fn wc_1840_table_5_2() {
    run_comparison_test("WC-1840", "WC/WC042-Table-5.docx", "WC/WC042-Table-5-Mod.docx", 2);
}

#[test]
fn wc_1850_nested_table() {
    run_comparison_test("WC-1850", "WC/WC043-Nested-Table.docx", "WC/WC043-Nested-Table-Mod.docx", 2);
}

// ============================================================================
// Text boxes
// ============================================================================

#[test]
fn wc_1860_text_box() {
    run_comparison_test("WC-1860", "WC/WC044-Text-Box.docx", "WC/WC044-Text-Box-Mod.docx", 2);
}

#[test]
fn wc_1870_text_box2() {
    run_comparison_test("WC-1870", "WC/WC045-Text-Box.docx", "WC/WC045-Text-Box-Mod.docx", 2);
}

#[test]
fn wc_1880_two_text_box() {
    run_comparison_test("WC-1880", "WC/WC046-Two-Text-Box.docx", "WC/WC046-Two-Text-Box-Mod.docx", 2);
}

#[test]
fn wc_1890_two_text_box2() {
    run_comparison_test("WC-1890", "WC/WC047-Two-Text-Box.docx", "WC/WC047-Two-Text-Box-Mod.docx", 2);
}

#[test]
fn wc_1900_text_box_in_cell() {
    run_comparison_test("WC-1900", "WC/WC048-Text-Box-in-Cell.docx", "WC/WC048-Text-Box-in-Cell-Mod.docx", 6);
}

#[test]
fn wc_1910_text_box_in_cell2() {
    run_comparison_test("WC-1910", "WC/WC049-Text-Box-in-Cell.docx", "WC/WC049-Text-Box-in-Cell-Mod.docx", 5);
}

#[test]
fn wc_1920_table_in_text_box() {
    run_comparison_test("WC-1920", "WC/WC050-Table-in-Text-Box.docx", "WC/WC050-Table-in-Text-Box-Mod.docx", 8);
}

#[test]
fn wc_1930_table_in_text_box2() {
    run_comparison_test("WC-1930", "WC/WC051-Table-in-Text-Box.docx", "WC/WC051-Table-in-Text-Box-Mod.docx", 9);
}

// ============================================================================
// SmartArt same
// ============================================================================

#[test]
fn wc_1940_smartart_same() {
    run_comparison_test("WC-1940", "WC/WC052-SmartArt-Same.docx", "WC/WC052-SmartArt-Same-Mod.docx", 2);
}

// ============================================================================
// Text in cell
// ============================================================================

#[test]
fn wc_1950_text_in_cell() {
    run_comparison_test("WC-1950", "WC/WC053-Text-in-Cell.docx", "WC/WC053-Text-in-Cell-Mod.docx", 2);
}

#[test]
fn wc_1960_text_in_cell_no_change() {
    run_comparison_test("WC-1960", "WC/WC054-Text-in-Cell.docx", "WC/WC054-Text-in-Cell-Mod.docx", 0);
}

// ============================================================================
// French language
// ============================================================================

#[test]
fn wc_1970_french() {
    run_comparison_test("WC-1970", "WC/WC055-French.docx", "WC/WC055-French-Mod.docx", 0);
}

#[test]
fn wc_1980_french2() {
    run_comparison_test("WC-1980", "WC/WC056-French.docx", "WC/WC056-French-Mod.docx", 0);
}

// ============================================================================
// Table merged cell
// ============================================================================

#[test]
fn wc_2000_table_merged_cell() {
    run_comparison_test("WC-2000", "WC/WC058-Table-Merged-Cell.docx", "WC/WC058-Table-Merged-Cell-Mod.docx", 6);
}

// ============================================================================
// More footnote/endnote tests
// ============================================================================

#[test]
fn wc_2010_footnote_complex() {
    run_comparison_test("WC-2010", "WC/WC059-Footnote.docx", "WC/WC059-Footnote-Mod.docx", 5);
}

#[test]
fn wc_2020_endnote_complex() {
    run_comparison_test("WC-2020", "WC/WC060-Endnote.docx", "WC/WC060-Endnote-Mod.docx", 3);
}

// ============================================================================
// Style added
// ============================================================================

#[test]
fn wc_2030_style_added() {
    run_comparison_test("WC-2030", "WC/WC061-Style-Added.docx", "WC/WC061-Style-Added-Mod.docx", 1);
}

// ============================================================================
// New character style
// ============================================================================

#[test]
fn wc_2040_new_char_style() {
    run_comparison_test("WC-2040", "WC/WC062-New-Char-Style-Added.docx", "WC/WC062-New-Char-Style-Added-Mod.docx", 3);
}

// ============================================================================
// More footnotes
// ============================================================================

#[test]
fn wc_2050_footnote() {
    run_comparison_test("WC-2050", "WC/WC063-Footnote.docx", "WC/WC063-Footnote-Mod.docx", 1);
}

#[test]
fn wc_2060_footnote_reverse() {
    run_comparison_test("WC-2060", "WC/WC063-Footnote-Mod.docx", "WC/WC063-Footnote.docx", 1);
}

#[test]
fn wc_2070_footnote_no_change() {
    run_comparison_test("WC-2070", "WC/WC064-Footnote.docx", "WC/WC064-Footnote-Mod.docx", 0);
}

// ============================================================================
// More textbox tests
// ============================================================================

#[test]
fn wc_2080_textbox() {
    run_comparison_test("WC-2080", "WC/WC065-Textbox.docx", "WC/WC065-Textbox-Mod.docx", 2);
}

#[test]
fn wc_2090_textbox_before_ins() {
    run_comparison_test("WC-2090", "WC/WC066-Textbox-Before-Ins.docx", "WC/WC066-Textbox-Before-Ins-Mod.docx", 1);
}

#[test]
fn wc_2092_textbox_before_ins_reverse() {
    run_comparison_test("WC-2092", "WC/WC066-Textbox-Before-Ins-Mod.docx", "WC/WC066-Textbox-Before-Ins.docx", 1);
}

#[test]
fn wc_2100_textbox_image() {
    run_comparison_test("WC-2100", "WC/WC067-Textbox-Image.docx", "WC/WC067-Textbox-Image-Mod.docx", 2);
}
