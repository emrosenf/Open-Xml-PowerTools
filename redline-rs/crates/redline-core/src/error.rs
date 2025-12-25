use thiserror::Error;

#[derive(Error, Debug)]
pub enum RedlineError {
    #[error("Invalid OOXML package: {message}")]
    InvalidPackage { message: String },

    #[error("Missing required part '{part_path}' in {document_type} document")]
    MissingPart { part_path: String, document_type: String },

    #[error("XML parsing error at {location}: {message}")]
    XmlParse { message: String, location: String },

    #[error("XML serialization error: {0}")]
    XmlWrite(String),

    #[error("Invalid relationship: {message}")]
    InvalidRelationship { message: String },

    #[error("Unsupported feature: {feature}")]
    UnsupportedFeature { feature: String },

    #[error("Comparison failed: {0}")]
    ComparisonFailed(String),

    #[error(transparent)]
    Io(#[from] std::io::Error),

    #[error(transparent)]
    Zip(#[from] zip::result::ZipError),
}

pub type Result<T> = std::result::Result<T, RedlineError>;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn error_display_formats_correctly() {
        let err = RedlineError::InvalidPackage {
            message: "test error".to_string(),
        };
        assert_eq!(err.to_string(), "Invalid OOXML package: test error");
    }

    #[test]
    fn error_missing_part_formats_correctly() {
        let err = RedlineError::MissingPart {
            part_path: "/word/document.xml".to_string(),
            document_type: "Word".to_string(),
        };
        assert_eq!(
            err.to_string(),
            "Missing required part '/word/document.xml' in Word document"
        );
    }
}
