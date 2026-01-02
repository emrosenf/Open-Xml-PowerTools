//! Build UI-ready change list from raw PmlChange data.
//!
//! This module transforms raw `PmlChange` records into `PmlChangeListItem`
//! records suitable for display in a UI change list sidebar.

use super::types::{
    PmlChange, PmlChangeDetails, PmlChangeListItem, PmlChangeListOptions, PmlChangeType,
};
#[allow(unused_imports)]
use super::types::PmlWordCount;

/// Build a UI-friendly change list from raw changes.
///
/// # Arguments
/// * `changes` - Raw change data from comparison
/// * `options` - Options controlling the transformation
///
/// # Returns
/// A vector of `PmlChangeListItem` records ready for UI display
pub fn build_change_list(
    changes: &[PmlChange],
    options: &PmlChangeListOptions,
) -> Vec<PmlChangeListItem> {
    let mut items = Vec::new();
    if changes.is_empty() {
        return items;
    }

    let mut current_group: Vec<&PmlChange> = Vec::new();

    for change in changes {
        if current_group.is_empty() {
            current_group.push(change);
            continue;
        }

        let last = current_group.last().unwrap();

        // Group by slide if option enabled
        let can_group = options.group_by_slide
            && change.slide_index == last.slide_index
            && change.slide_index.is_some()
            && is_groupable_change(change.change_type)
            && is_groupable_change(last.change_type);

        if can_group {
            current_group.push(change);
        } else {
            // Flush group
            flush_group(&current_group, &mut items, options);
            current_group.clear();
            current_group.push(change);
        }
    }

    if !current_group.is_empty() {
        flush_group(&current_group, &mut items, options);
    }

    items
}

fn is_groupable_change(t: PmlChangeType) -> bool {
    // Only group structural/content changes on the same slide
    !matches!(
        t,
        PmlChangeType::SlideInserted
            | PmlChangeType::SlideDeleted
            | PmlChangeType::SlideMoved
            | PmlChangeType::SlideLayoutChanged
            | PmlChangeType::SlideBackgroundChanged
            | PmlChangeType::SlideNotesChanged
    )
}

fn flush_group(
    group: &[&PmlChange],
    items: &mut Vec<PmlChangeListItem>,
    options: &PmlChangeListOptions,
) {
    if group.is_empty() {
        return;
    }

    // If grouping is enabled and we have multiple items, try to consolidate
    if options.group_by_slide && group.len() > 1 {
        // Group by shape within the slide group
        let mut shape_groups: std::collections::HashMap<String, Vec<&PmlChange>> =
            std::collections::HashMap::new();
        let mut ungrouped = Vec::new();

        for c in group {
            if let Some(shape_name) = &c.shape_name {
                shape_groups.entry(shape_name.clone()).or_default().push(c);
            } else {
                ungrouped.push(c);
            }
        }

        // Process shape groups
        for (shape_name, shape_changes) in shape_groups {
            if shape_changes.len() > 1 {
                items.push(create_grouped_item(
                    &shape_changes,
                    items.len() + 1,
                    Some(&shape_name),
                ));
            } else {
                items.push(create_single_item(
                    shape_changes[0],
                    items.len() + 1,
                    options,
                ));
            }
        }

        // Process ungrouped
        for c in ungrouped {
            items.push(create_single_item(c, items.len() + 1, options));
        }
    } else {
        // Just add them individually
        for c in group {
            items.push(create_single_item(c, items.len() + 1, options));
        }
    }
}

fn create_single_item(
    change: &PmlChange,
    id_suffix: usize,
    options: &PmlChangeListOptions,
) -> PmlChangeListItem {
    let summary = change.get_description();

    // Preview text for text changes
    let preview_text = if let Some(text_changes) = &change.text_changes {
        if !text_changes.is_empty() {
            // Use new text or old text from first text change
            let text = text_changes[0]
                .new_text
                .as_deref()
                .or(text_changes[0].old_text.as_deref())
                .unwrap_or("");
            Some(truncate_text(text, options.max_preview_length))
        } else {
            None
        }
    } else {
        change
            .new_value
            .as_deref()
            .or(change.old_value.as_deref())
            .map(|s| truncate_text(s, options.max_preview_length))
    };

    PmlChangeListItem {
        id: format!("change-{}", id_suffix),
        change_type: change.change_type,
        slide_index: change.slide_index,
        shape_name: change.shape_name.clone(),
        shape_id: change.shape_id.clone(),
        summary,
        preview_text,
        word_count: None, // TODO: Calculate word count if needed
        count: None,
        details: Some(PmlChangeDetails {
            old_value: change.old_value.clone(),
            new_value: change.new_value.clone(),
            old_slide_index: change.old_slide_index,
            text_changes: change.text_changes.clone(),
            match_confidence: change.match_confidence,
        }),
        anchor: build_anchor(change),
    }
}

fn create_grouped_item(
    group: &[&PmlChange],
    id_suffix: usize,
    shape_name: Option<&str>,
) -> PmlChangeListItem {
    let first = group[0];
    let count = group.len();

    let summary = format!(
        "{} changes in '{}' on slide {}",
        count,
        shape_name.unwrap_or("Shape"),
        first.slide_index.unwrap_or(0)
    );

    PmlChangeListItem {
        id: format!("change-{}", id_suffix),
        change_type: first.change_type, // Use first change type as representative
        slide_index: first.slide_index,
        shape_name: first.shape_name.clone(),
        shape_id: first.shape_id.clone(),
        summary,
        preview_text: None,
        word_count: None,
        count: Some(count),
        details: None, // No details for grouped item
        anchor: build_anchor(first),
    }
}

fn build_anchor(change: &PmlChange) -> Option<String> {
    // Format: slide-X-shape-ID
    if let Some(slide) = change.slide_index {
        if let Some(shape_id) = &change.shape_id {
            return Some(format!("slide-{}-shape-{}", slide, shape_id));
        }
        return Some(format!("slide-{}", slide));
    }
    None
}

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
    fn test_build_change_list_groups_shape_changes() {
        let changes = vec![
            PmlChange {
                change_type: PmlChangeType::TextChanged,
                slide_index: Some(1),
                shape_name: Some("Shape1".to_string()),
                shape_id: Some("1".to_string()),
                ..Default::default()
            },
            PmlChange {
                change_type: PmlChangeType::ShapeMoved,
                slide_index: Some(1),
                shape_name: Some("Shape1".to_string()),
                shape_id: Some("1".to_string()),
                ..Default::default()
            },
        ];

        let options = PmlChangeListOptions::default();
        let items = build_change_list(&changes, &options);

        assert_eq!(items.len(), 1);
        assert_eq!(items[0].count, Some(2));
        assert!(items[0].summary.contains("2 changes in 'Shape1'"));
    }

    #[test]
    fn test_build_change_list_slide_separation() {
        let changes = vec![
            PmlChange {
                change_type: PmlChangeType::TextChanged,
                slide_index: Some(1),
                shape_name: Some("Shape1".to_string()),
                ..Default::default()
            },
            PmlChange {
                change_type: PmlChangeType::TextChanged,
                slide_index: Some(2), // Different slide
                shape_name: Some("Shape1".to_string()),
                ..Default::default()
            },
        ];

        let options = PmlChangeListOptions::default();
        let items = build_change_list(&changes, &options);

        assert_eq!(items.len(), 2);
    }
}
