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
    
    // Fuzzy matching settings
    pub enable_fuzzy_shape_matching: bool,
    pub slide_similarity_threshold: f64,
    pub shape_similarity_threshold: f64,
    
    // Tolerance settings
    pub position_tolerance: i64, // EMUs
    
    // Legacy/Other tolerance settings (keeping for compatibility if used elsewhere)
    pub transform_tolerance: TransformTolerance,
    
    // === Output Settings ===
    /// Author name for change annotations
    pub author_for_changes: String,
    
    /// Add a summary slide at the end of marked presentations
    pub add_summary_slide: bool,
    
    /// Add change summary to speaker notes
    pub add_notes_annotations: bool,
    
    // === Colors (RRGGBB hex) ===
    /// Color for inserted elements
    pub inserted_color: String,
    
    /// Color for deleted elements
    pub deleted_color: String,
    
    /// Color for modified elements
    pub modified_color: String,
    
    /// Color for moved elements
    pub moved_color: String,
    
    /// Color for formatting-only changes
    pub formatting_color: String,
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
            
            enable_fuzzy_shape_matching: true,
            slide_similarity_threshold: 0.4,
            shape_similarity_threshold: 0.7,
            position_tolerance: 91440, // ~0.1 inch
            
            transform_tolerance: TransformTolerance::default(),
            
            // Output settings
            author_for_changes: "Open-Xml-PowerTools".to_string(),
            add_summary_slide: true,
            add_notes_annotations: true,
            
            // Colors
            inserted_color: "00AA00".to_string(),  // Green
            deleted_color: "FF0000".to_string(),   // Red
            modified_color: "FFA500".to_string(),  // Orange
            moved_color: "0000FF".to_string(),     // Blue
            formatting_color: "9932CC".to_string(), // Purple
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
        assert_eq!(settings.slide_similarity_threshold, 0.4);
        assert_eq!(settings.position_tolerance, 91440);
    }
}
