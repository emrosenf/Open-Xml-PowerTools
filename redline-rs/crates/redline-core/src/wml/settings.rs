use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WmlComparerSettings {
    pub word_separators: Vec<char>,
    pub case_insensitive: bool,
    pub conflate_breaking_and_nonbreaking_spaces: bool,
    pub culture_info: Option<String>,
    pub author_for_revisions: String,
    pub date_time_for_revisions: DateTime<Utc>,
    pub track_formatting_changes: bool,
    pub detail_threshold: f64,
    pub starting_id_for_footnotes_endnotes: i32,
}

impl Default for WmlComparerSettings {
    fn default() -> Self {
        Self {
            word_separators: vec![
                ' ', '-', ')', '(', ';', ',', '.', '!', '?', ':', '\'', '"',
                '/', '\\', '[', ']', '{', '}', '<', '>',
                '\u{FF08}', '\u{FF09}', '\u{FF0C}', '\u{3001}', '\u{FF0C}',
                '\u{FF1B}', '\u{3002}', '\u{FF1A}', '\u{7684}',
            ],
            case_insensitive: false,
            conflate_breaking_and_nonbreaking_spaces: true,
            culture_info: None,
            author_for_revisions: "Open-Xml-PowerTools".to_string(),
            date_time_for_revisions: Utc::now(),
            track_formatting_changes: true,
            detail_threshold: 0.15,
            starting_id_for_footnotes_endnotes: 1,
        }
    }
}

impl WmlComparerSettings {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn with_author(mut self, author: &str) -> Self {
        self.author_for_revisions = author.to_string();
        self
    }

    pub fn with_case_insensitive(mut self, case_insensitive: bool) -> Self {
        self.case_insensitive = case_insensitive;
        self
    }

    pub fn with_track_formatting(mut self, track: bool) -> Self {
        self.track_formatting_changes = track;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_settings_have_expected_values() {
        let settings = WmlComparerSettings::default();
        
        assert!(!settings.case_insensitive);
        assert!(settings.conflate_breaking_and_nonbreaking_spaces);
        assert!(settings.track_formatting_changes);
        assert!((settings.detail_threshold - 0.15).abs() < f64::EPSILON);
        assert!(settings.word_separators.contains(&' '));
        assert!(settings.word_separators.contains(&'-'));
    }

    #[test]
    fn builder_pattern_works() {
        let settings = WmlComparerSettings::new()
            .with_author("Test Author")
            .with_case_insensitive(true)
            .with_track_formatting(false);
        
        assert_eq!(settings.author_for_revisions, "Test Author");
        assert!(settings.case_insensitive);
        assert!(!settings.track_formatting_changes);
    }
}
