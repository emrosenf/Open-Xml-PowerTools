//! OOXML Validation Module
//!
//! Validates Office Open XML documents for structural correctness and semantic rules.
//!
//! ## Key OOXML Rules Enforced:
//!
//! 1. **Element Ordering**: Certain elements must appear in specific positions:
//!    - `<w:pPr>` must be the FIRST child of `<w:p>` (paragraph properties)
//!    - `<w:rPr>` must be the FIRST child of `<w:r>` (run properties)
//!    - `<w:tblPr>` must be the FIRST child of `<w:tbl>` (table properties)
//!    - `<w:trPr>` must be the FIRST child of `<w:tr>` (table row properties)
//!    - `<w:tcPr>` must be the FIRST child of `<w:tc>` (table cell properties)
//!
//! 2. **Required Parts**: Core OOXML structure requirements:
//!    - `[Content_Types].xml` must exist
//!    - `_rels/.rels` must exist
//!
//! 3. **XML Well-formedness**: All XML parts must be well-formed.

use std::io::Read;
use roxmltree::{Document, Node};

#[derive(Debug, Clone)]
pub struct ValidationResult {
    pub is_valid: bool,
    pub errors: Vec<ValidationError>,
    pub warnings: Vec<ValidationWarning>,
}

#[derive(Debug, Clone)]
pub struct ValidationError {
    pub path: String,
    pub message: String,
    pub error_type: ValidationErrorType,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ValidationErrorType {
    MissingPart,
    InvalidXml,
    BrokenRelationship,
    InvalidContentType,
    SchemaViolation,
    /// Element ordering violation (e.g., pPr not first child of p)
    ElementOrderingViolation,
}

#[derive(Debug, Clone)]
pub struct ValidationWarning {
    pub path: String,
    pub message: String,
}

/// WordprocessingML namespace
const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// Validate an OOXML document from bytes
pub fn validate_ooxml(doc_bytes: &[u8]) -> ValidationResult {
    let mut errors = Vec::new();
    let mut warnings = Vec::new();

    let cursor = std::io::Cursor::new(doc_bytes);
    let mut archive = match zip::ZipArchive::new(cursor) {
        Ok(a) => a,
        Err(e) => {
            errors.push(ValidationError {
                path: String::new(),
                message: format!("Invalid ZIP archive: {}", e),
                error_type: ValidationErrorType::InvalidXml,
            });
            return ValidationResult {
                is_valid: false,
                errors,
                warnings,
            };
        }
    };

    let file_names: Vec<_> = (0..archive.len())
        .filter_map(|i| archive.by_index(i).ok())
        .map(|f| f.name().to_string())
        .collect();

    // Check required parts
    if !file_names.iter().any(|n| n == "[Content_Types].xml") {
        errors.push(ValidationError {
            path: "[Content_Types].xml".to_string(),
            message: "Missing [Content_Types].xml".to_string(),
            error_type: ValidationErrorType::MissingPart,
        });
    }

    if !file_names.iter().any(|n| n == "_rels/.rels") {
        errors.push(ValidationError {
            path: "_rels/.rels".to_string(),
            message: "Missing _rels/.rels".to_string(),
            error_type: ValidationErrorType::MissingPart,
        });
    }

    // Validate WML parts for element ordering
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
        "word/comments.xml",
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

    ValidationResult {
        is_valid: errors.is_empty(),
        errors,
        warnings,
    }
}

/// Validate WML element ordering rules
///
/// OOXML ISO 29500 specifies strict ordering for child elements within many container elements.
/// This function checks the most critical ordering rules that affect document rendering.
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

    // Check all elements recursively
    check_element_ordering(&doc.root(), part_name, &mut errors);

    errors
}

fn check_element_ordering(node: &Node, part_name: &str, errors: &mut Vec<ValidationError>) {
    if node.is_element() {
        let ns = node.tag_name().namespace();
        let local = node.tag_name().name();

        // Only check elements in the W namespace
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

    // Recursively check children
    for child in node.children() {
        check_element_ordering(&child, part_name, errors);
    }
}

/// Check that if a properties element exists, it must be the first element child
fn check_first_child_rule(
    parent: &Node,
    props_local_name: &str,
    parent_description: &str,
    part_name: &str,
    errors: &mut Vec<ValidationError>,
) {
    let element_children: Vec<_> = parent
        .children()
        .filter(|c| c.is_element())
        .collect();

    if element_children.is_empty() {
        return;
    }

    // Find if the properties element exists anywhere in children
    let props_position = element_children.iter().position(|c| {
        c.tag_name().namespace() == Some(W_NS) && c.tag_name().name() == props_local_name
    });

    // If properties element exists but is not first, that's an error
    if let Some(pos) = props_position {
        if pos != 0 {
            // Get some context about where this is in the document
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

/// Validate a WML document and return detailed results
pub fn validate_wml_document(doc_bytes: &[u8]) -> ValidationResult {
    validate_ooxml(doc_bytes)
}

/// Assert that a document has no OOXML validation errors
/// Panics with detailed error messages if validation fails
pub fn assert_valid_ooxml(doc_bytes: &[u8], context: &str) {
    let result = validate_ooxml(doc_bytes);
    if !result.is_valid {
        let error_messages: Vec<_> = result
            .errors
            .iter()
            .map(|e| format!("  - [{}] {}: {}", format!("{:?}", e.error_type), e.path, e.message))
            .collect();
        panic!(
            "OOXML validation failed for {}:\n{}",
            context,
            error_messages.join("\n")
        );
    }
}

/// Check only element ordering rules (faster, focused check)
pub fn validate_element_ordering_only(doc_bytes: &[u8]) -> ValidationResult {
    let mut errors = Vec::new();
    let warnings = Vec::new();

    let cursor = std::io::Cursor::new(doc_bytes);
    let mut archive = match zip::ZipArchive::new(cursor) {
        Ok(a) => a,
        Err(e) => {
            errors.push(ValidationError {
                path: String::new(),
                message: format!("Invalid ZIP archive: {}", e),
                error_type: ValidationErrorType::InvalidXml,
            });
            return ValidationResult {
                is_valid: false,
                errors,
                warnings,
            };
        }
    };

    let file_names: Vec<_> = (0..archive.len())
        .filter_map(|i| archive.by_index(i).ok())
        .map(|f| f.name().to_string())
        .collect();

    // Only check element ordering in WML parts
    let wml_parts = [
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
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

    ValidationResult {
        is_valid: errors.is_empty(),
        errors,
        warnings,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn validate_detects_invalid_zip() {
        let result = validate_ooxml(b"not a zip file");
        assert!(!result.is_valid);
        assert!(!result.errors.is_empty());
    }

    #[test]
    fn validate_ppr_must_be_first_child_of_p() {
        // Valid: pPr is first child
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

        // Invalid: pPr is NOT first child (run comes before it)
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
        assert!(errors[0].message.contains("pPr"));
    }

    #[test]
    fn validate_rpr_must_be_first_child_of_r() {
        // Valid: rPr is first child of run
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

        // Invalid: rPr comes after text
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
        assert!(errors[0].message.contains("rPr"));
    }

    #[test]
    fn validate_table_properties_ordering() {
        // Valid: tblPr, trPr, tcPr all first in their containers
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

        // Invalid: tcPr comes after paragraph
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
        assert!(errors[0].message.contains("tcPr"));
    }

    #[test]
    fn validate_paragraph_without_ppr_is_valid() {
        // Paragraphs don't require pPr - they just need it FIRST if present
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

    #[test]
    fn validate_empty_paragraph_is_valid() {
        let xml = r#"<?xml version="1.0"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:body>
                    <w:p/>
                </w:body>
            </w:document>"#;

        let errors = validate_wml_element_ordering("word/document.xml", xml);
        assert!(errors.is_empty(), "Empty paragraph should be valid: {:?}", errors);
    }
}
