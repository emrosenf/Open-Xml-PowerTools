use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PmlComparerSettings {
    pub author_for_revisions: String,
    pub date_time_for_revisions: DateTime<Utc>,
    pub compare_slide_structure: bool,
    pub compare_shape_structure: bool,
    pub compare_text_content: bool,
    pub compare_text_formatting: bool,
    pub compare_shape_transforms: bool,
    pub use_slide_alignment_lcs: bool,
    pub transform_tolerance: TransformTolerance,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TransformTolerance {
    pub position: f64,
    pub size: f64,
    pub rotation: f64,
}

impl Default for TransformTolerance {
    fn default() -> Self {
        Self {
            position: 0.01,
            size: 0.01,
            rotation: 0.01,
        }
    }
}

impl Default for PmlComparerSettings {
    fn default() -> Self {
        Self {
            author_for_revisions: "Open-Xml-PowerTools".to_string(),
            date_time_for_revisions: Utc::now(),
            compare_slide_structure: true,
            compare_shape_structure: true,
            compare_text_content: true,
            compare_text_formatting: true,
            compare_shape_transforms: true,
            use_slide_alignment_lcs: true,
            transform_tolerance: TransformTolerance::default(),
        }
    }
}

impl PmlComparerSettings {
    pub fn new() -> Self {
        Self::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_settings_have_expected_values() {
        let settings = PmlComparerSettings::default();
        
        assert!(settings.compare_slide_structure);
        assert!(settings.compare_shape_structure);
        assert!(settings.compare_text_content);
        assert!(settings.use_slide_alignment_lcs);
    }
}
