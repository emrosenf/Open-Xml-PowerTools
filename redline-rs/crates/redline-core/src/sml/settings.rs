use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlComparerSettings {
    pub author_for_revisions: String,
    pub date_time_for_revisions: DateTime<Utc>,
    pub numeric_tolerance: Option<f64>,
    pub ignore_whitespace: bool,
    pub case_sensitive: bool,
    pub compare_formatting: bool,
    pub enable_row_alignment: bool,
    pub enable_column_alignment: bool,
    pub sheet_rename_threshold: f64,
}

impl Default for SmlComparerSettings {
    fn default() -> Self {
        Self {
            author_for_revisions: "Open-Xml-PowerTools".to_string(),
            date_time_for_revisions: Utc::now(),
            numeric_tolerance: None,
            ignore_whitespace: true,
            case_sensitive: true,
            compare_formatting: true,
            enable_row_alignment: true,
            enable_column_alignment: true,
            sheet_rename_threshold: 0.7,
        }
    }
}

impl SmlComparerSettings {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn with_numeric_tolerance(mut self, tolerance: f64) -> Self {
        self.numeric_tolerance = Some(tolerance);
        self
    }

    pub fn with_formatting(mut self, compare: bool) -> Self {
        self.compare_formatting = compare;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_settings_have_expected_values() {
        let settings = SmlComparerSettings::default();
        
        assert!(settings.case_sensitive);
        assert!(settings.compare_formatting);
        assert!(settings.enable_row_alignment);
        assert!((settings.sheet_rename_threshold - 0.7).abs() < f64::EPSILON);
    }
}
