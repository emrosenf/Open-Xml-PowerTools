//! WML (Word) document comparison types
//!
//! Defines types and enums for comparing Word documents with
//! detailed change tracking for UI display.

use serde::{Deserialize, Serialize};

/// Types of changes detected during Word document comparison.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub enum WmlChangeType {
    /// Text content inserted
    TextInserted,
    /// Text content deleted
    TextDeleted,
    /// Text content replaced (delete + insert at same location)
    TextReplaced,
    /// Entire paragraph inserted
    ParagraphInserted,
    /// Entire paragraph deleted
    ParagraphDeleted,
    /// Text formatting changed (bold, italic, etc.)
    FormatChanged,
    /// Table row inserted
    TableRowInserted,
    /// Table row deleted
    TableRowDeleted,
    /// Table cell content changed
    TableCellChanged,
    /// Content moved from one location
    MovedFrom,
    /// Content moved to new location
    MovedTo,
    /// Footnote/endnote changed
    NoteChanged,
    /// Image/drawing inserted
    ImageInserted,
    /// Image/drawing deleted
    ImageDeleted,
    /// Image/drawing replaced
    ImageReplaced,
}

/// Word count statistics for a change.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WmlWordCount {
    /// Number of words deleted
    pub deleted: usize,
    /// Number of words inserted
    pub inserted: usize,
}

/// Represents a single change between two Word documents.
/// This is the raw change data captured during comparison.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WmlChange {
    /// Type of change
    pub change_type: WmlChangeType,
    /// Unique revision ID from the Word markup (w:id)
    pub revision_id: i32,
    /// Zero-based index of the paragraph containing this change
    pub paragraph_index: Option<usize>,
    /// For table changes, the row index
    pub table_row_index: Option<usize>,
    /// For table changes, the cell index
    pub table_cell_index: Option<usize>,
    /// Original text (for deletions and replacements)
    pub old_text: Option<String>,
    /// New text (for insertions and replacements)
    pub new_text: Option<String>,
    /// Word count statistics
    pub word_count: Option<WmlWordCount>,
    /// For format changes, description of what changed
    pub format_description: Option<String>,
    /// Author who made the change
    pub author: Option<String>,
    /// Date/time of the change (ISO 8601)
    pub date_time: Option<String>,
    /// Whether this is inside a footnote
    pub in_footnote: bool,
    /// Whether this is inside an endnote
    pub in_endnote: bool,
    /// Whether this is inside a table
    pub in_table: bool,
    /// Whether this is inside a textbox
    pub in_textbox: bool,
}

impl Default for WmlChange {
    fn default() -> Self {
        Self {
            change_type: WmlChangeType::TextInserted,
            revision_id: 0,
            paragraph_index: None,
            table_row_index: None,
            table_cell_index: None,
            old_text: None,
            new_text: None,
            word_count: None,
            format_description: None,
            author: None,
            date_time: None,
            in_footnote: false,
            in_endnote: false,
            in_table: false,
            in_textbox: false,
        }
    }
}

/// Result of a Word document comparison with detailed change tracking.
#[derive(Debug, Clone)]
pub struct WmlComparisonResult {
    /// The comparison result document bytes
    pub document: Vec<u8>,
    /// List of all individual changes detected
    pub changes: Vec<WmlChange>,
    /// Number of insertions
    pub insertions: usize,
    /// Number of deletions
    pub deletions: usize,
    /// Number of format changes
    pub format_changes: usize,
    /// Total number of revisions
    pub revision_count: usize,
}

/// UI-friendly representation of a change for display in a change list.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WmlChangeListItem {
    /// Unique identifier for this change list item
    pub id: String,
    /// Type of change
    pub change_type: WmlChangeType,
    /// Human-readable summary of the change
    pub summary: String,
    /// Preview text showing what changed
    pub preview_text: Option<String>,
    /// Word count statistics
    pub word_count: Option<WmlWordCount>,
    /// Zero-based paragraph index for navigation
    pub paragraph_index: Option<usize>,
    /// Revision ID for navigation to the change in the document
    pub revision_id: Option<i32>,
    /// Anchor string for navigation (e.g., "para-5" or "revision-12")
    pub anchor: Option<String>,
    /// Additional details about the change
    pub details: Option<WmlChangeDetails>,
}

/// Additional details about a change for display.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WmlChangeDetails {
    /// Original text before change
    pub old_text: Option<String>,
    /// New text after change
    pub new_text: Option<String>,
    /// Format change description
    pub format_description: Option<String>,
    /// Author of the change
    pub author: Option<String>,
    /// Date/time of change
    pub date_time: Option<String>,
    /// Location context (e.g., "In footnote", "In table row 2")
    pub location_context: Option<String>,
}

/// Options for building a change list from comparison results.
#[derive(Debug, Clone)]
pub struct WmlChangeListOptions {
    /// Whether to group adjacent changes of the same type.
    /// Default: true
    pub group_adjacent_changes: bool,
    /// Whether to merge delete+insert pairs into "replaced" changes.
    /// Default: true
    pub merge_replacements: bool,
    /// Maximum length for preview text before truncation.
    /// Default: 100
    pub max_preview_length: usize,
}

impl Default for WmlChangeListOptions {
    fn default() -> Self {
        Self {
            group_adjacent_changes: true,
            merge_replacements: true,
            max_preview_length: 100,
        }
    }
}

/// Revision count statistics
#[derive(Debug, Clone, Default)]
pub struct RevisionCounts {
    pub insertions: usize,
    pub deletions: usize,
    pub format_changes: usize,
}

impl RevisionCounts {
    pub fn total(&self) -> usize {
        self.insertions + self.deletions + self.format_changes
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn change_type_serializes_correctly() {
        let change_type = WmlChangeType::TextInserted;
        let json = serde_json::to_string(&change_type).unwrap();
        assert_eq!(json, "\"TextInserted\"");
    }

    #[test]
    fn wml_change_default() {
        let change = WmlChange::default();
        assert_eq!(change.change_type, WmlChangeType::TextInserted);
        assert_eq!(change.revision_id, 0);
        assert!(!change.in_table);
    }

    #[test]
    fn change_list_options_default() {
        let options = WmlChangeListOptions::default();
        assert!(options.group_adjacent_changes);
        assert!(options.merge_replacements);
        assert_eq!(options.max_preview_length, 100);
    }
}
