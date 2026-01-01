pub mod normalizer;
pub mod validator;

pub use normalizer::Normalizer;
pub use validator::{
    validate_ooxml,
    validate_element_ordering_only,
    validate_wml_element_ordering,
    assert_valid_ooxml,
    ValidationResult,
    ValidationError,
    ValidationErrorType,
};
