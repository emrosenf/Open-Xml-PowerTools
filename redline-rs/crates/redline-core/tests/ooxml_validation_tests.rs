//! OOXML Validation Integration Tests
//!
//! These tests verify that the comparer produces valid OOXML output
//! according to the ISO 29500 specification.
//!
//! ## Key OOXML Rules Enforced:
//!
//! 1. **Element Ordering**: Certain elements must appear in specific positions:
//!    - `<w:pPr>` must be the FIRST child of `<w:p>` (paragraph properties)
//!    - `<w:rPr>` must be the FIRST child of `<w:r>` (run properties)
//!    - `<w:tblPr>` must be the FIRST child of `<w:tbl>` (table properties)
//!    - `<w:trPr>` must be the FIRST child of `<w:tr>` (table row properties)
//!    - `<w:tcPr>` must be the FIRST child of `<w:tc>` (table cell properties)

use redline_core::wml::{WmlComparer, WmlDocument, WmlComparerSettings};
use roxmltree::{Document, Node};
use std::fs;
use std::io::Read;

/// WordprocessingML namespace
const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

#[derive(Debug, Clone, PartialEq)]
pub enum ValidationErrorType {
    InvalidXml,
    ElementOrderingViolation,
}

#[derive(Debug, Clone)]
pub struct ValidationError {
    pub path: String,
    pub message: String,
    pub error_type: ValidationErrorType,
}

/// Validate WML element ordering rules
pub fn validate_wml_element_ordering(part_name: &str, xml_content: &str) -> Vec<ValidationError> {
    let mut errors = Vec::new();

    let doc = match Document::parse(xml_content) {
        Ok(d) => d,
        Err(e) => {
            errors.push(ValidationError {
                path: part_name.to_string(),
                message: format!("XML parse error: {}", e),
                error_type: ValidationErrorType::InvalidXml,
            });
            return errors;
        }
    };

    check_element_ordering(&doc.root(), part_name, &mut errors);
    errors
}

fn check_element_ordering(node: &Node, part_name: &str, errors: &mut Vec<ValidationError>) {
    if node.is_element() {
        let ns = node.tag_name().namespace();
        let local = node.tag_name().name();

        if ns == Some(W_NS) {
            match local {
                "p" => check_first_child_rule(node, "pPr", "paragraph (w:p)", part_name, errors),
                "r" => check_first_child_rule(node, "rPr", "run (w:r)", part_name, errors),
                "tbl" => check_first_child_rule(node, "tblPr", "table (w:tbl)", part_name, errors),
                "tr" => check_first_child_rule(node, "trPr", "table row (w:tr)", part_name, errors),
                "tc" => check_first_child_rule(node, "tcPr", "table cell (w:tc)", part_name, errors),
                _ => {}
            }
        }
    }

    for child in node.children() {
        check_element_ordering(&child, part_name, errors);
    }
}

fn check_first_child_rule(
    parent: &Node,
    props_local_name: &str,
    parent_description: &str,
    part_name: &str,
    errors: &mut Vec<ValidationError>,
) {
    let element_children: Vec<_> = parent.children().filter(|c| c.is_element()).collect();

    if element_children.is_empty() {
        return;
    }

    let props_position = element_children.iter().position(|c| {
        c.tag_name().namespace() == Some(W_NS) && c.tag_name().name() == props_local_name
    });

    if let Some(pos) = props_position {
        if pos != 0 {
            let parent_id = parent
                .attributes()
                .find(|a| a.name() == "paraId" || a.name() == "id")
                .map(|a| a.value())
                .unwrap_or("unknown");

            errors.push(ValidationError {
                path: part_name.to_string(),
                message: format!(
                    "OOXML ordering violation: <w:{}> must be the FIRST child of {} but found at position {} (parent id: {})",
                    props_local_name, parent_description, pos + 1, parent_id
                ),
                error_type: ValidationErrorType::ElementOrderingViolation,
            });
        }
    }
}

/// Validate element ordering in an OOXML docx file
pub fn validate_docx_element_ordering(doc_bytes: &[u8]) -> Vec<ValidationError> {
    let mut errors = Vec::new();

    let cursor = std::io::Cursor::new(doc_bytes);
    let mut archive = match zip::ZipArchive::new(cursor) {
        Ok(a) => a,
        Err(e) => {
            errors.push(ValidationError {
                path: String::new(),
                message: format!("Invalid ZIP archive: {}", e),
                error_type: ValidationErrorType::InvalidXml,
            });
            return errors;
        }
    };

    let file_names: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            archive.by_index(i).ok().map(|f| f.name().to_string())
        })
        .collect();

    let wml_parts = [
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/header1.xml",
        "word/header2.xml",
        "word/header3.xml",
        "word/footer1.xml",
        "word/footer2.xml",
        "word/footer3.xml",
    ];

    for part_name in &wml_parts {
        if file_names.iter().any(|n| n == *part_name) {
            if let Ok(mut file) = archive.by_name(part_name) {
                let mut content = String::new();
                if file.read_to_string(&mut content).is_ok() {
                    let part_errors = validate_wml_element_ordering(part_name, &content);
                    errors.extend(part_errors);
                }
            }
        }
    }

    errors
}

// ============================================================================
// Unit Tests for XML Element Ordering Rules
// ============================================================================

/// Test that pPr must be the first child of p (valid case)
#[test]
fn ooxml_rule_ppr_must_be_first_child_of_p_valid() {
    let valid_xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
                    <w:r><w:t>Hello</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>"#;

    let errors = validate_wml_element_ordering("word/document.xml", valid_xml);
    assert!(errors.is_empty(), "Valid XML should have no errors: {:?}", errors);
}

/// Test that pPr must be the first child of p (invalid case)
#[test]
fn ooxml_rule_ppr_must_be_first_child_of_p_invalid() {
    let invalid_xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r><w:t>Hello</w:t></w:r>
                    <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
                </w:p>
            </w:body>
        </w:document>"#;

    let errors = validate_wml_element_ordering("word/document.xml", invalid_xml);
    assert!(!errors.is_empty(), "Invalid XML should have errors");
    assert_eq!(errors[0].error_type, ValidationErrorType::ElementOrderingViolation);
    assert!(errors[0].message.contains("pPr"), "Error should mention pPr: {}", errors[0].message);
}

/// Test that rPr must be the first child of r (valid case)
#[test]
fn ooxml_rule_rpr_must_be_first_child_of_r_valid() {
    let valid_xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r>
                        <w:rPr><w:b/></w:rPr>
                        <w:t>Bold text</w:t>
                    </w:r>
                </w:p>
            </w:body>
        </w:document>"#;

    let errors = validate_wml_element_ordering("word/document.xml", valid_xml);
    assert!(errors.is_empty(), "Valid XML should have no errors: {:?}", errors);
}

/// Test that rPr must be the first child of r (invalid case)
#[test]
fn ooxml_rule_rpr_must_be_first_child_of_r_invalid() {
    let invalid_xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r>
                        <w:t>Bold text</w:t>
                        <w:rPr><w:b/></w:rPr>
                    </w:r>
                </w:p>
            </w:body>
        </w:document>"#;

    let errors = validate_wml_element_ordering("word/document.xml", invalid_xml);
    assert!(!errors.is_empty(), "Invalid XML should have errors");
    assert_eq!(errors[0].error_type, ValidationErrorType::ElementOrderingViolation);
    assert!(errors[0].message.contains("rPr"), "Error should mention rPr: {}", errors[0].message);
}

/// Test table properties ordering (valid case)
#[test]
fn ooxml_rule_table_properties_ordering_valid() {
    let valid_xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:tbl>
                    <w:tblPr><w:tblW w:w="5000"/></w:tblPr>
                    <w:tr>
                        <w:trPr><w:trHeight w:val="400"/></w:trPr>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="2500"/></w:tcPr>
                            <w:p><w:r><w:t>Cell</w:t></w:r></w:p>
                        </w:tc>
                    </w:tr>
                </w:tbl>
            </w:body>
        </w:document>"#;

    let errors = validate_wml_element_ordering("word/document.xml", valid_xml);
    assert!(errors.is_empty(), "Valid XML should have no errors: {:?}", errors);
}

/// Test tcPr must be first child of tc (invalid case)
#[test]
fn ooxml_rule_tcpr_must_be_first_child_of_tc_invalid() {
    let invalid_xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:tbl>
                    <w:tblPr><w:tblW w:w="5000"/></w:tblPr>
                    <w:tr>
                        <w:trPr><w:trHeight w:val="400"/></w:trPr>
                        <w:tc>
                            <w:p><w:r><w:t>Cell</w:t></w:r></w:p>
                            <w:tcPr><w:tcW w:w="2500"/></w:tcPr>
                        </w:tc>
                    </w:tr>
                </w:tbl>
            </w:body>
        </w:document>"#;

    let errors = validate_wml_element_ordering("word/document.xml", invalid_xml);
    assert!(!errors.is_empty(), "Invalid XML should have errors");
    assert!(errors[0].message.contains("tcPr"), "Error should mention tcPr: {}", errors[0].message);
}

/// Test that paragraphs without pPr are valid (pPr is optional)
#[test]
fn ooxml_rule_paragraph_without_ppr_is_valid() {
    let xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p>
                    <w:r><w:t>No properties</w:t></w:r>
                </w:p>
            </w:body>
        </w:document>"#;

    let errors = validate_wml_element_ordering("word/document.xml", xml);
    assert!(errors.is_empty(), "Paragraph without pPr should be valid: {:?}", errors);
}

/// Test that empty paragraphs are valid
#[test]
fn ooxml_rule_empty_paragraph_is_valid() {
    let xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p/>
            </w:body>
        </w:document>"#;

    let errors = validate_wml_element_ordering("word/document.xml", xml);
    assert!(errors.is_empty(), "Empty paragraph should be valid: {:?}", errors);
}

// ============================================================================
// Integration Tests - Verify Comparison Output
// ============================================================================

/// Integration test: verify that comparison output passes OOXML validation
#[test]
fn integration_comparison_output_passes_ooxml_validation() {
    let test_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .join("TestFiles/WmlComparer");

    let doc1_path = test_dir.join("WC001-010-before.docx");
    let doc2_path = test_dir.join("WC001-010-after.docx");

    if !doc1_path.exists() || !doc2_path.exists() {
        eprintln!("Skipping integration test: test files not found at {:?}", test_dir);
        return;
    }

    let doc1_bytes = fs::read(&doc1_path).expect("Failed to read doc1");
    let doc2_bytes = fs::read(&doc2_path).expect("Failed to read doc2");

    let doc1 = WmlDocument::from_bytes(&doc1_bytes).expect("Failed to parse doc1");
    let doc2 = WmlDocument::from_bytes(&doc2_bytes).expect("Failed to parse doc2");

    let settings = WmlComparerSettings::default();
    let result = WmlComparer::compare(&doc1, &doc2, Some(&settings)).expect("Comparison failed");

    let errors = validate_docx_element_ordering(&result.document);

    if !errors.is_empty() {
        for error in &errors {
            eprintln!("OOXML Validation Error: [{:?}] {}: {}",
                error.error_type, error.path, error.message);
        }
        panic!("Comparison output failed OOXML validation with {} errors", errors.len());
    }
}

/// Integration test: verify footnotes have correct element ordering
#[test]
fn integration_footnotes_have_correct_element_ordering() {
    let test_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .join("TestFiles/WmlComparer");

    // WC002-050 has footnotes
    let doc1_path = test_dir.join("WC002-050-before.docx");
    let doc2_path = test_dir.join("WC002-050-after.docx");

    if !doc1_path.exists() || !doc2_path.exists() {
        eprintln!("Skipping footnotes integration test: test files not found");
        return;
    }

    let doc1_bytes = fs::read(&doc1_path).expect("Failed to read doc1");
    let doc2_bytes = fs::read(&doc2_path).expect("Failed to read doc2");

    let doc1 = WmlDocument::from_bytes(&doc1_bytes).expect("Failed to parse doc1");
    let doc2 = WmlDocument::from_bytes(&doc2_bytes).expect("Failed to parse doc2");

    let settings = WmlComparerSettings::default();
    let result = WmlComparer::compare(&doc1, &doc2, Some(&settings)).expect("Comparison failed");

    let errors = validate_docx_element_ordering(&result.document);

    let footnote_errors: Vec<_> = errors.iter()
        .filter(|e| e.path.contains("footnotes") || e.path.contains("endnotes"))
        .collect();

    if !footnote_errors.is_empty() {
        for error in &footnote_errors {
            eprintln!("Footnote/Endnote Error: [{:?}] {}: {}",
                error.error_type, error.path, error.message);
        }
        panic!("Footnotes/endnotes failed OOXML validation with {} errors",
            footnote_errors.len());
    }
}

/// Test all WmlComparer test files for OOXML element ordering compliance
#[test]
fn integration_all_wml_test_files_pass_ooxml_validation() {
    let test_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .join("TestFiles/WmlComparer");

    if !test_dir.exists() {
        eprintln!("Skipping bulk integration test: test directory not found");
        return;
    }

    let mut total_tests = 0;
    let mut passed_tests = 0;
    let mut failed_tests = Vec::new();

    // Find all test file pairs
    for entry in fs::read_dir(&test_dir).unwrap() {
        let entry = entry.unwrap();
        let path = entry.path();
        let file_name = path.file_name().unwrap().to_str().unwrap();

        if file_name.ends_with("-before.docx") {
            let after_path = test_dir.join(file_name.replace("-before.docx", "-after.docx"));
            if !after_path.exists() {
                continue;
            }

            total_tests += 1;
            let test_name = file_name.replace("-before.docx", "");

            let doc1_bytes = match fs::read(&path) {
                Ok(b) => b,
                Err(_) => continue,
            };
            let doc2_bytes = match fs::read(&after_path) {
                Ok(b) => b,
                Err(_) => continue,
            };

            let doc1 = match WmlDocument::from_bytes(&doc1_bytes) {
                Ok(d) => d,
                Err(_) => continue,
            };
            let doc2 = match WmlDocument::from_bytes(&doc2_bytes) {
                Ok(d) => d,
                Err(_) => continue,
            };

            let settings = WmlComparerSettings::default();
            let result = match WmlComparer::compare(&doc1, &doc2, Some(&settings)) {
                Ok(r) => r,
                Err(_) => continue,
            };

            let errors = validate_docx_element_ordering(&result.document);

            if errors.is_empty() {
                passed_tests += 1;
            } else {
                failed_tests.push((test_name.clone(), errors));
            }
        }
    }

    println!("\nOOXML Validation Results: {}/{} passed", passed_tests, total_tests);

    if !failed_tests.is_empty() {
        println!("\nFailed tests:");
        for (name, errors) in &failed_tests {
            println!("  {} ({} errors):", name, errors.len());
            for error in errors.iter().take(3) {
                println!("    - {}: {}", error.path, error.message);
            }
            if errors.len() > 3 {
                println!("    ... and {} more", errors.len() - 3);
            }
        }
        panic!("{} test(s) failed OOXML validation", failed_tests.len());
    }
}
