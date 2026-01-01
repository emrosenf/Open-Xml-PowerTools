//! Build UI-ready change list from raw WmlChange data.
//!
//! This module transforms raw `WmlChange` records into `WmlChangeListItem`
//! records suitable for display in a UI change list sidebar.

use super::types::{
    WmlChange, WmlChangeDetails, WmlChangeListItem, WmlChangeListOptions, WmlChangeType,
    WmlWordCount,
};

/// Build a UI-friendly change list from raw changes.
///
/// This function transforms raw `WmlChange` records into `WmlChangeListItem`
/// records with the following transformations:
///
/// 1. **Merge replacements**: Adjacent delete+insert pairs with the same
///    paragraph index are merged into a single "Replaced" item.
/// 2. **Truncate previews**: Preview text is truncated to `max_preview_length`.
/// 3. **Generate summaries**: Human-readable summary strings are generated.
/// 4. **Build anchors**: Navigation anchors using revision IDs are created.
/// 5. **Location context**: Context strings like "In table" are added.
///
/// # Arguments
/// * `changes` - Raw change data from `extract_changes_from_document`
/// * `options` - Options controlling the transformation
///
/// # Returns
/// A vector of `WmlChangeListItem` records ready for UI display
pub fn build_change_list(
    changes: &[WmlChange],
    options: &WmlChangeListOptions,
) -> Vec<WmlChangeListItem> {
    let mut items = Vec::new();
    let mut i = 0;

    while i < changes.len() {
        let change = &changes[i];
        let next_change = changes.get(i + 1);

        // Check for delete+insert replacement pattern
        if options.merge_replacements
            && change.change_type == WmlChangeType::TextDeleted
            && next_change.map(|c| c.change_type) == Some(WmlChangeType::TextInserted)
            && change.paragraph_index == next_change.unwrap().paragraph_index
        {
            let next = next_change.unwrap();
            let deleted_words = change.word_count.as_ref().map(|w| w.deleted).unwrap_or(0);
            let inserted_words = next.word_count.as_ref().map(|w| w.inserted).unwrap_or(0);

            let old_text = change.old_text.as_deref().unwrap_or("");
            let new_text = next.new_text.as_deref().unwrap_or("");

            let preview = format!(
                "{} → {}",
                truncate_text(old_text, options.max_preview_length / 2),
                truncate_text(new_text, options.max_preview_length / 2)
            );

            items.push(WmlChangeListItem {
                id: format!("change-{}", items.len() + 1),
                change_type: WmlChangeType::TextReplaced,
                summary: "Replaced".to_string(),
                preview_text: Some(preview),
                word_count: Some(WmlWordCount {
                    deleted: deleted_words,
                    inserted: inserted_words,
                }),
                paragraph_index: change.paragraph_index,
                revision_id: Some(change.revision_id),
                anchor: Some(format!("revision-{}", change.revision_id)),
                details: Some(WmlChangeDetails {
                    old_text: change.old_text.clone(),
                    new_text: next.new_text.clone(),
                    format_description: None,
                    author: change.author.clone(),
                    date_time: change.date_time.clone(),
                    location_context: build_location_context(change),
                }),
            });

            i += 2; // Skip both the delete and insert
            continue;
        }

        // Convert single change to list item
        items.push(to_change_list_item(
            change,
            items.len() + 1,
            options.max_preview_length,
        ));
        i += 1;
    }

    items
}

/// Convert a single WmlChange to a WmlChangeListItem
fn to_change_list_item(
    change: &WmlChange,
    index: usize,
    max_preview_length: usize,
) -> WmlChangeListItem {
    let summary = summarize_change(change);
    let preview_text = change
        .new_text
        .as_deref()
        .or(change.old_text.as_deref())
        .unwrap_or("");

    WmlChangeListItem {
        id: format!("change-{}", index),
        change_type: change.change_type,
        summary,
        preview_text: Some(truncate_text(preview_text, max_preview_length)),
        word_count: change.word_count.clone(),
        paragraph_index: change.paragraph_index,
        revision_id: Some(change.revision_id),
        anchor: Some(format!("revision-{}", change.revision_id)),
        details: Some(WmlChangeDetails {
            old_text: change.old_text.clone(),
            new_text: change.new_text.clone(),
            format_description: change.format_description.clone(),
            author: change.author.clone(),
            date_time: change.date_time.clone(),
            location_context: build_location_context(change),
        }),
    }
}

/// Generate a human-readable summary for a change
fn summarize_change(change: &WmlChange) -> String {
    match change.change_type {
        WmlChangeType::TextInserted => "Inserted".to_string(),
        WmlChangeType::TextDeleted => "Deleted".to_string(),
        WmlChangeType::TextReplaced => "Replaced".to_string(),
        WmlChangeType::ParagraphInserted => "Paragraph inserted".to_string(),
        WmlChangeType::ParagraphDeleted => "Paragraph deleted".to_string(),
        WmlChangeType::FormatChanged => "Format changed".to_string(),
        WmlChangeType::TableRowInserted => "Table row inserted".to_string(),
        WmlChangeType::TableRowDeleted => "Table row deleted".to_string(),
        WmlChangeType::TableCellChanged => "Table cell changed".to_string(),
        WmlChangeType::ImageInserted => "Image inserted".to_string(),
        WmlChangeType::ImageDeleted => "Image deleted".to_string(),
        WmlChangeType::ImageReplaced => "Image replaced".to_string(),
        WmlChangeType::NoteChanged => "Note changed".to_string(),
        WmlChangeType::MovedFrom => "Moved from".to_string(),
        WmlChangeType::MovedTo => "Moved to".to_string(),
    }
}

/// Build a location context string from change flags
fn build_location_context(change: &WmlChange) -> Option<String> {
    let mut parts = Vec::new();

    if change.in_footnote {
        parts.push("In footnote");
    }
    if change.in_endnote {
        parts.push("In endnote");
    }
    if change.in_table {
        parts.push("In table");
    }
    if change.in_textbox {
        parts.push("In textbox");
    }

    if parts.is_empty() {
        None
    } else {
        Some(parts.join(", "))
    }
}

/// Truncate text to a maximum length, adding "..." if truncated
fn truncate_text(text: &str, max_length: usize) -> String {
    if text.len() <= max_length {
        text.to_string()
    } else if max_length <= 3 {
        "...".to_string()
    } else {
        format!("{}...", &text[..max_length - 3])
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_truncate_text() {
        assert_eq!(truncate_text("hello", 10), "hello");
        assert_eq!(truncate_text("hello world", 8), "hello...");
        assert_eq!(truncate_text("hi", 2), "hi");
        assert_eq!(truncate_text("hello", 3), "...");
    }

    #[test]
    fn test_summarize_change() {
        let mut change = WmlChange::default();
        change.change_type = WmlChangeType::TextInserted;
        assert_eq!(summarize_change(&change), "Inserted");

        change.change_type = WmlChangeType::TextDeleted;
        assert_eq!(summarize_change(&change), "Deleted");

        change.change_type = WmlChangeType::FormatChanged;
        assert_eq!(summarize_change(&change), "Format changed");
    }

    #[test]
    fn test_build_location_context() {
        let mut change = WmlChange::default();
        assert_eq!(build_location_context(&change), None);

        change.in_table = true;
        assert_eq!(
            build_location_context(&change),
            Some("In table".to_string())
        );

        change.in_footnote = true;
        assert_eq!(
            build_location_context(&change),
            Some("In footnote, In table".to_string())
        );
    }

    #[test]
    fn test_build_change_list_merges_replacements() {
        let changes = vec![
            WmlChange {
                change_type: WmlChangeType::TextDeleted,
                revision_id: 1,
                paragraph_index: Some(1),
                old_text: Some("old text".to_string()),
                word_count: Some(WmlWordCount {
                    deleted: 2,
                    inserted: 0,
                }),
                ..Default::default()
            },
            WmlChange {
                change_type: WmlChangeType::TextInserted,
                revision_id: 2,
                paragraph_index: Some(1),
                new_text: Some("new text".to_string()),
                word_count: Some(WmlWordCount {
                    deleted: 0,
                    inserted: 2,
                }),
                ..Default::default()
            },
        ];

        let options = WmlChangeListOptions::default();
        let items = build_change_list(&changes, &options);

        assert_eq!(items.len(), 1);
        assert_eq!(items[0].change_type, WmlChangeType::TextReplaced);
        assert_eq!(items[0].summary, "Replaced");
        assert!(items[0].preview_text.as_ref().unwrap().contains("→"));
    }

    #[test]
    fn test_build_change_list_no_merge_different_paragraphs() {
        let changes = vec![
            WmlChange {
                change_type: WmlChangeType::TextDeleted,
                revision_id: 1,
                paragraph_index: Some(1),
                old_text: Some("old text".to_string()),
                ..Default::default()
            },
            WmlChange {
                change_type: WmlChangeType::TextInserted,
                revision_id: 2,
                paragraph_index: Some(2), // Different paragraph!
                new_text: Some("new text".to_string()),
                ..Default::default()
            },
        ];

        let options = WmlChangeListOptions::default();
        let items = build_change_list(&changes, &options);

        // Should NOT merge - different paragraphs
        assert_eq!(items.len(), 2);
        assert_eq!(items[0].change_type, WmlChangeType::TextDeleted);
        assert_eq!(items[1].change_type, WmlChangeType::TextInserted);
    }
}
