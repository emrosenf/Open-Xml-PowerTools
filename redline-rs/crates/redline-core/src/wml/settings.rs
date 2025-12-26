use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};

/// Settings for WmlComparer document comparison.
/// Faithful port of WmlComparerSettings from C# OpenXmlPowerTools.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WmlComparerSettings {
    /// Characters that separate words for comparison purposes.
    /// Default includes space, punctuation, and Chinese punctuation.
    pub word_separators: Vec<char>,
    
    /// Author name for tracked revisions. If None, the author will be extracted
    /// from the modified document's LastModifiedBy or Creator core property.
    pub author_for_revisions: Option<String>,
    
    /// Date/time for tracked revisions in ISO 8601 format.
    /// Defaults to current time when created.
    pub date_time_for_revisions: String,
    
    /// Threshold for detail level in comparison (0.0-1.0).
    /// Lower values provide more detailed comparison.
    pub detail_threshold: f64,
    
    /// Whether to perform case-insensitive comparison.
    pub case_insensitive: bool,
    
    /// Whether to treat breaking and non-breaking spaces as equivalent.
    pub conflate_breaking_and_nonbreaking_spaces: bool,
    
    /// Whether to track formatting changes as revisions.
    pub track_formatting_changes: bool,
    
    /// Culture info for locale-specific comparison (e.g., "en-US", "zh-CN").
    pub culture_info: Option<String>,
    
    /// Starting ID for footnote and endnote numbering.
    pub starting_id_for_footnotes_endnotes: i32,
}

impl Default for WmlComparerSettings {
    fn default() -> Self {
        Self {
            // C# default: new[] { ' ', '-', ')', '(', ';', ',', '（', '）', '，', '、', '、', '，', '；', '。', '：', '的', }
            word_separators: vec![
                ' ', '-', ')', '(', ';', ',',
                '（', // U+FF08 FULLWIDTH LEFT PARENTHESIS
                '）', // U+FF09 FULLWIDTH RIGHT PARENTHESIS
                '，', // U+FF0C FULLWIDTH COMMA
                '、', // U+3001 IDEOGRAPHIC COMMA
                '、', // U+3001 IDEOGRAPHIC COMMA (duplicate in C# source)
                '，', // U+FF0C FULLWIDTH COMMA (duplicate in C# source)
                '；', // U+FF1B FULLWIDTH SEMICOLON
                '。', // U+3002 IDEOGRAPHIC FULL STOP
                '：', // U+FF1A FULLWIDTH COLON
                '的', // U+7684 CJK UNIFIED IDEOGRAPH (Chinese possessive particle)
            ],
            author_for_revisions: None, // C# default: null
            date_time_for_revisions: Utc::now().to_rfc3339(), // C# default: DateTime.Now.ToString("o")
            detail_threshold: 0.15, // C# default: 0.15
            case_insensitive: false, // C# default: false
            conflate_breaking_and_nonbreaking_spaces: true, // C# default: true
            track_formatting_changes: true, // C# default: true
            culture_info: None, // C# default: null
            starting_id_for_footnotes_endnotes: 1, // C# default: 1
        }
    }
}

impl WmlComparerSettings {
    /// Create new settings with default values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Builder method to set author for revisions.
    pub fn with_author(mut self, author: impl Into<String>) -> Self {
        self.author_for_revisions = Some(author.into());
        self
    }

    /// Builder method to set case sensitivity.
    pub fn with_case_insensitive(mut self, case_insensitive: bool) -> Self {
        self.case_insensitive = case_insensitive;
        self
    }

    /// Builder method to set formatting tracking.
    pub fn with_track_formatting(mut self, track: bool) -> Self {
        self.track_formatting_changes = track;
        self
    }
    
    /// Builder method to set culture info.
    pub fn with_culture_info(mut self, culture: impl Into<String>) -> Self {
        self.culture_info = Some(culture.into());
        self
    }
    
    /// Check if a character is a word separator.
    /// C# equivalent: IsWordSeparator(char c)
    pub fn is_word_separator(&self, c: char) -> bool {
        self.word_separators.contains(&c)
    }
}

/// Settings for consolidating multiple document revisions.
/// Faithful port of WmlComparerConsolidateSettings from C# OpenXmlPowerTools.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WmlComparerConsolidateSettings {
    /// Whether to consolidate revisions with a table.
    pub consolidate_with_table: bool,
}

impl Default for WmlComparerConsolidateSettings {
    fn default() -> Self {
        Self {
            consolidate_with_table: true, // C# default: true
        }
    }
}

/// Information about a revised document.
/// Faithful port of WmlRevisedDocumentInfo from C# OpenXmlPowerTools.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WmlRevisedDocumentInfo {
    /// The revised document.
    pub revised_document: Vec<u8>, // Placeholder for WmlDocument type
    
    /// Name of the revisor.
    pub revisor: String,
    
    /// Color associated with this revision (RGB).
    pub color: (u8, u8, u8),
}

/// Type of revision in a compared document.
/// Faithful port of WmlComparerRevisionType enum from C# OpenXmlPowerTools.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum WmlComparerRevisionType {
    /// Content was inserted.
    Inserted,
    
    /// Content was deleted.
    Deleted,
    
    /// Formatting was changed.
    FormatChanged,
}

/// A single revision extracted from a compared document.
/// Faithful port of WmlComparerRevision class from C# OpenXmlPowerTools.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WmlComparerRevision {
    /// Type of this revision.
    pub revision_type: WmlComparerRevisionType,
    
    /// Text content of the revision.
    pub text: String,
    
    /// Author who made the revision.
    pub author: String,
    
    /// Date when the revision was made.
    pub date: String,
    
    /// XML element containing the content (serialized).
    /// C# type: XElement
    pub content_x_element: Option<String>,
    
    /// XML element representing the revision markup (serialized).
    /// C# type: XElement
    pub revision_x_element: Option<String>,
    
    /// URI of the part containing this revision.
    /// C# type: Uri
    pub part_uri: Option<String>,
    
    /// Content type of the part.
    pub part_content_type: Option<String>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_settings_have_expected_values() {
        let settings = WmlComparerSettings::default();
        
        // Check boolean defaults
        assert!(!settings.case_insensitive);
        assert!(settings.conflate_breaking_and_nonbreaking_spaces);
        assert!(settings.track_formatting_changes);
        
        // Check numeric defaults
        assert!((settings.detail_threshold - 0.15).abs() < f64::EPSILON);
        assert_eq!(settings.starting_id_for_footnotes_endnotes, 1);
        
        // Check Option defaults
        assert!(settings.author_for_revisions.is_none());
        assert!(settings.culture_info.is_none());
        
        // Check word separators contain expected characters
        assert!(settings.word_separators.contains(&' '));
        assert!(settings.word_separators.contains(&'-'));
        assert!(settings.word_separators.contains(&'（')); // Chinese parenthesis
        assert!(settings.word_separators.contains(&'的')); // Chinese particle
        
        // Verify exact length matches C# (16 chars including duplicates)
        assert_eq!(settings.word_separators.len(), 16);
    }

    #[test]
    fn builder_pattern_works() {
        let settings = WmlComparerSettings::new()
            .with_author("Test Author")
            .with_case_insensitive(true)
            .with_track_formatting(false)
            .with_culture_info("en-US");
        
        assert_eq!(settings.author_for_revisions, Some("Test Author".to_string()));
        assert!(settings.case_insensitive);
        assert!(!settings.track_formatting_changes);
        assert_eq!(settings.culture_info, Some("en-US".to_string()));
    }
    
    #[test]
    fn is_word_separator_works() {
        let settings = WmlComparerSettings::default();
        
        assert!(settings.is_word_separator(' '));
        assert!(settings.is_word_separator('-'));
        assert!(settings.is_word_separator('（'));
        assert!(!settings.is_word_separator('a'));
        assert!(!settings.is_word_separator('Z'));
    }
    
    #[test]
    fn consolidate_settings_defaults() {
        let settings = WmlComparerConsolidateSettings::default();
        assert!(settings.consolidate_with_table);
    }
    
    #[test]
    fn revision_type_enum_values() {
        use WmlComparerRevisionType::*;
        
        let inserted = Inserted;
        let deleted = Deleted;
        let format_changed = FormatChanged;
        
        assert_ne!(inserted, deleted);
        assert_ne!(deleted, format_changed);
        assert_ne!(format_changed, inserted);
    }
}
