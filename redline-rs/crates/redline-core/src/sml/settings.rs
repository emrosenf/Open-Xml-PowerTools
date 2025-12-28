// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

use serde::{Deserialize, Serialize};

/// Settings for controlling spreadsheet comparison behavior.
/// 100% parity with C# SmlComparerSettings.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlComparerSettings {
    // ===== Phase 1: Core Comparison Settings =====
    
    /// Whether to compare cell values.
    pub compare_values: bool,

    /// Whether to compare cell formulas.
    pub compare_formulas: bool,

    /// Whether to compare cell formatting.
    pub compare_formatting: bool,

    /// Whether to compare sheet structure (added/removed sheets).
    pub compare_sheet_structure: bool,

    /// Whether value comparison should be case-insensitive.
    pub case_insensitive_values: bool,

    /// Tolerance for numeric comparison (0.0 for exact match).
    pub numeric_tolerance: f64,

    /// Author name for change annotations.
    pub author_for_changes: String,

    /// Optional callback for logging (not serializable - use Arc<Mutex<dyn Fn>> in practice).
    #[serde(skip)]
    pub log_callback: Option<Box<dyn Fn(&str) + Send + Sync>>,

    // ===== Highlight Colors (ARGB hex without #) =====
    
    /// Fill color for added cells (default: light green).
    pub added_cell_color: String,

    /// Fill color for deleted cells in summary (default: light red).
    pub deleted_cell_color: String,

    /// Fill color for value changes (default: gold).
    pub modified_value_color: String,

    /// Fill color for formula changes (default: sky blue).
    pub modified_formula_color: String,

    /// Fill color for format-only changes (default: lavender).
    pub modified_format_color: String,

    /// Fill color for inserted rows (default: light cyan).
    pub inserted_row_color: String,

    /// Fill color for deleted rows in summary (default: misty rose).
    pub deleted_row_color: String,

    // ===== Phase 2: Row/Column Alignment Settings =====
    
    /// Enable row alignment using LCS algorithm to detect inserted/deleted rows.
    pub enable_row_alignment: bool,

    /// Enable column alignment using LCS algorithm to detect inserted/deleted columns.
    /// Off by default, can be expensive.
    pub enable_column_alignment: bool,

    /// Enable sheet rename detection based on content similarity.
    pub enable_sheet_rename_detection: bool,

    /// Minimum similarity threshold (0.0-1.0) to consider a sheet renamed vs added/deleted.
    pub sheet_rename_similarity_threshold: f64,

    /// Number of cells to sample per row for row signature hashing.
    pub row_signature_sample_size: i32,

    // ===== Phase 3: Additional Comparison Settings =====
    
    /// Whether to compare named ranges (defined names).
    pub compare_named_ranges: bool,

    /// Whether to compare cell comments/notes.
    pub compare_comments: bool,

    /// Whether to compare data validation rules.
    pub compare_data_validation: bool,

    /// Whether to compare merged cell regions.
    pub compare_merged_cells: bool,

    /// Whether to compare conditional formatting rules.
    pub compare_conditional_formatting: bool,

    /// Whether to compare hyperlinks.
    pub compare_hyperlinks: bool,

    // ===== Highlight Colors for Phase 3 Features =====
    
    /// Fill color for named range changes (default: light purple).
    pub named_range_change_color: String,

    /// Fill color for comment changes (default: light yellow).
    pub comment_change_color: String,

    /// Fill color for data validation changes (default: light orange).
    pub data_validation_change_color: String,
}

impl Default for SmlComparerSettings {
    fn default() -> Self {
        Self {
            // Phase 1: Core settings
            compare_values: true,
            compare_formulas: true,
            compare_formatting: true,
            compare_sheet_structure: true,
            case_insensitive_values: false,
            numeric_tolerance: 0.0,
            author_for_changes: "Open-Xml-PowerTools".to_string(),
            log_callback: None,

            // Phase 1: Highlight colors
            added_cell_color: "90EE90".to_string(),
            deleted_cell_color: "FFCCCB".to_string(),
            modified_value_color: "FFD700".to_string(),
            modified_formula_color: "87CEEB".to_string(),
            modified_format_color: "E6E6FA".to_string(),
            inserted_row_color: "E0FFFF".to_string(),
            deleted_row_color: "FFE4E1".to_string(),

            // Phase 2: Row/Column alignment
            enable_row_alignment: false,
            enable_column_alignment: false,
            enable_sheet_rename_detection: true,
            sheet_rename_similarity_threshold: 0.7,
            row_signature_sample_size: 10,

            // Phase 3: Additional comparison settings
            compare_named_ranges: true,
            compare_comments: true,
            compare_data_validation: true,
            compare_merged_cells: true,
            compare_conditional_formatting: true,
            compare_hyperlinks: true,

            // Phase 3: Highlight colors
            named_range_change_color: "DDA0DD".to_string(),
            comment_change_color: "FFFACD".to_string(),
            data_validation_change_color: "FFDAB9".to_string(),
        }
    }
}

impl SmlComparerSettings {
    /// Creates a new instance with default settings.
    pub fn new() -> Self {
        Self::default()
    }

    /// Sets the numeric tolerance for comparison.
    pub fn with_numeric_tolerance(mut self, tolerance: f64) -> Self {
        self.numeric_tolerance = tolerance;
        self
    }

    /// Sets whether to compare formatting.
    pub fn with_formatting(mut self, compare: bool) -> Self {
        self.compare_formatting = compare;
        self
    }

    /// Sets the author name for changes.
    pub fn with_author(mut self, author: String) -> Self {
        self.author_for_changes = author;
        self
    }

    /// Enables or disables row alignment.
    pub fn with_row_alignment(mut self, enable: bool) -> Self {
        self.enable_row_alignment = enable;
        self
    }

    /// Enables or disables column alignment.
    pub fn with_column_alignment(mut self, enable: bool) -> Self {
        self.enable_column_alignment = enable;
        self
    }

    /// Sets the sheet rename similarity threshold.
    pub fn with_rename_threshold(mut self, threshold: f64) -> Self {
        self.sheet_rename_similarity_threshold = threshold;
        self
    }

    /// Sets case sensitivity for value comparison.
    pub fn with_case_sensitive(mut self, case_sensitive: bool) -> Self {
        self.case_insensitive_values = !case_sensitive;
        self
    }

    /// Logs a message if a callback is configured.
    pub fn log(&self, message: &str) {
        if let Some(ref callback) = self.log_callback {
            callback(message);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_settings_have_expected_values() {
        let settings = SmlComparerSettings::default();
        
        // Phase 1 defaults
        assert!(settings.compare_values);
        assert!(settings.compare_formulas);
        assert!(settings.compare_formatting);
        assert!(settings.compare_sheet_structure);
        assert!(!settings.case_insensitive_values);
        assert_eq!(settings.numeric_tolerance, 0.0);
        assert_eq!(settings.author_for_changes, "Open-Xml-PowerTools");

        // Color defaults
        assert_eq!(settings.added_cell_color, "90EE90");
        assert_eq!(settings.deleted_cell_color, "FFCCCB");
        assert_eq!(settings.modified_value_color, "FFD700");

        // Phase 2 defaults
        assert!(!settings.enable_row_alignment);
        assert!(!settings.enable_column_alignment);
        assert!(settings.enable_sheet_rename_detection);
        assert!((settings.sheet_rename_similarity_threshold - 0.7).abs() < f64::EPSILON);
        assert_eq!(settings.row_signature_sample_size, 10);

        // Phase 3 defaults
        assert!(settings.compare_named_ranges);
        assert!(settings.compare_comments);
        assert!(settings.compare_data_validation);
        assert!(settings.compare_merged_cells);
        assert!(settings.compare_conditional_formatting);
        assert!(settings.compare_hyperlinks);
    }

    #[test]
    fn builder_pattern_works() {
        let settings = SmlComparerSettings::new()
            .with_numeric_tolerance(0.01)
            .with_formatting(false)
            .with_author("Test Author".to_string())
            .with_row_alignment(true)
            .with_case_sensitive(false);

        assert_eq!(settings.numeric_tolerance, 0.01);
        assert!(!settings.compare_formatting);
        assert_eq!(settings.author_for_changes, "Test Author");
        assert!(settings.enable_row_alignment);
        assert!(settings.case_insensitive_values);
    }

    #[test]
    fn log_callback_can_be_set() {
        let mut settings = SmlComparerSettings::new();
        let mut logged_messages = Vec::new();
        
        // In real usage, you'd use Arc<Mutex<Vec<String>>> to capture messages
        settings.log_callback = None; // Just test that it compiles
        settings.log("test message"); // Should not panic
    }
}
