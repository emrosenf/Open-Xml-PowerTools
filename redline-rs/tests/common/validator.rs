use std::collections::HashMap;

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

#[derive(Debug, Clone)]
pub enum ValidationErrorType {
    MissingPart,
    InvalidXml,
    BrokenRelationship,
    InvalidContentType,
    SchemaViolation,
}

#[derive(Debug, Clone)]
pub struct ValidationWarning {
    pub path: String,
    pub message: String,
}

pub fn validate_ooxml(doc_bytes: &[u8]) -> ValidationResult {
    let mut errors = Vec::new();
    let mut warnings = Vec::new();

    let cursor = std::io::Cursor::new(doc_bytes);
    let archive = match zip::ZipArchive::new(cursor) {
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
}
