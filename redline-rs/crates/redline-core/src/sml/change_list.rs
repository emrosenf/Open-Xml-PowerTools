//! Build UI-ready change list from raw SmlChange data.
//!
//! This module transforms raw `SmlChange` records into `SmlChangeListItem`
//! records suitable for display in a UI change list sidebar.

use super::types::{
    SmlChange, SmlChangeDetails, SmlChangeListItem, SmlChangeListOptions, SmlChangeType,
};

/// Build a UI-friendly change list from raw changes.
///
/// # Arguments
/// * `changes` - Raw change data from comparison
/// * `options` - Options controlling the transformation
///
/// # Returns
/// A vector of `SmlChangeListItem` records ready for UI display
pub fn build_change_list(
    changes: &[SmlChange],
    options: &SmlChangeListOptions,
) -> Vec<SmlChangeListItem> {
    let mut items = Vec::new();
    if changes.is_empty() {
        return items;
    }

    let mut current_group: Vec<&SmlChange> = Vec::new();
    // 0 = Unknown, 1 = Horizontal (Same Row), 2 = Vertical (Same Col)
    let mut group_direction = 0;

    for change in changes {
        if current_group.is_empty() {
            current_group.push(change);
            group_direction = 0;
            continue;
        }

        let last = current_group.last().unwrap();
        let adjacency = check_adjacency(last, change);

        let can_group = options.group_adjacent_cells
            && is_cell_change(change.change_type)
            && change.change_type == last.change_type
            && change.sheet_name == last.sheet_name
            && adjacency != 0
            && (group_direction == 0 || group_direction == adjacency);

        if can_group {
            current_group.push(change);
            if group_direction == 0 {
                group_direction = adjacency;
            }
        } else {
            // Flush group
            items.push(create_list_item(&current_group, items.len() + 1));
            current_group.clear();
            current_group.push(change);
            group_direction = 0;
        }
    }

    if !current_group.is_empty() {
        items.push(create_list_item(&current_group, items.len() + 1));
    }

    items
}

fn is_cell_change(t: SmlChangeType) -> bool {
    matches!(
        t,
        SmlChangeType::CellAdded
            | SmlChangeType::CellDeleted
            | SmlChangeType::ValueChanged
            | SmlChangeType::FormulaChanged
            | SmlChangeType::FormatChanged
    )
}

/// Check adjacency between two changes.
/// Returns: 0 = Not adjacent, 1 = Horizontal, 2 = Vertical
fn check_adjacency(c1: &SmlChange, c2: &SmlChange) -> i32 {
    // Must compare indices if available
    // Note: row_index and column_index might be populated for structural changes
    // For cell changes, we might need to parse cell_address if indices aren't guaranteed?
    // SmlDiffEngine logic usually populates SmlChange, let's see.
    // The current SmlChange definition has row_index/col_index as optional.
    // SmlDiffEngine uses indices for Row/Col changes, but for Cell changes it calls add_change with cell_address.
    // Wait, SmlDiffEngine does NOT populate row_index/column_index for Cell changes in the code I read.
    // It only populates cell_address.
    // So I need to parse cell_address here to check adjacency.

    let (r1, k1) = parse_cell_ref(c1.cell_address.as_deref().unwrap_or(""));
    let (r2, k2) = parse_cell_ref(c2.cell_address.as_deref().unwrap_or(""));

    if r1 == 0 || r2 == 0 {
        return 0; // Invalid or missing address
    }

    if r1 == r2 && (k1 as i32 - k2 as i32).abs() == 1 {
        return 1; // Horizontal
    }

    if k1 == k2 && (r1 as i32 - r2 as i32).abs() == 1 {
        return 2; // Vertical
    }

    0
}

/// Parse "A1" into (row, col) 1-based indices.
/// Very basic parser.
fn parse_cell_ref(address: &str) -> (u32, u32) {
    // Assuming standard "A1" format
    let mut chars = address.chars().peekable();
    let mut col_str = String::new();

    while let Some(c) = chars.peek() {
        if c.is_alphabetic() {
            col_str.push(*c);
            chars.next();
        } else {
            break;
        }
    }

    let row_str: String = chars.collect();

    let row = row_str.parse::<u32>().unwrap_or(0);
    let col = column_name_to_index(&col_str);

    (row, col)
}

fn column_name_to_index(name: &str) -> u32 {
    let mut index = 0;
    for c in name.chars() {
        if c.is_ascii_uppercase() {
            index = index * 26 + (c as u32 - 'A' as u32 + 1);
        } else if c.is_ascii_lowercase() {
            index = index * 26 + (c as u32 - 'a' as u32 + 1);
        }
    }
    index
}

fn create_list_item(group: &[&SmlChange], id_suffix: usize) -> SmlChangeListItem {
    let first = group[0];
    let count = group.len();

    let summary = if count == 1 {
        first.get_description()
    } else {
        format!(
            "{} cells changed ({}) in {}",
            count,
            summarize_type(first.change_type),
            first.sheet_name.as_deref().unwrap_or("Sheet")
        )
    };

    let details = if count == 1 {
        Some(SmlChangeDetails {
            old_value: first.old_value.clone(),
            new_value: first.new_value.clone(),
            old_formula: first.old_formula.clone(),
            new_formula: first.new_formula.clone(),
            old_format: first.old_format.clone(),
            new_format: first.new_format.clone(),
            old_comment: first.old_comment.clone(),
            new_comment: first.new_comment.clone(),
            comment_author: first.comment_author.clone(),
            data_validation_type: first.data_validation_type.clone(),
            old_data_validation: first.old_data_validation.clone(),
            new_data_validation: first.new_data_validation.clone(),
            merged_cell_range: first.merged_cell_range.clone(),
            old_hyperlink: first.old_hyperlink.clone(),
            new_hyperlink: first.new_hyperlink.clone(),
            old_sheet_name: first.old_sheet_name.clone(),
            new_sheet_name: None, // usually not relevant for details view unless rename
        })
    } else {
        None // No details for grouped items (or maybe summary details?)
    };

    let anchor = if let (Some(sheet), Some(addr)) = (&first.sheet_name, &first.cell_address) {
        Some(format!("{}!{}", sheet, addr))
    } else {
        None
    };

    SmlChangeListItem {
        id: format!("change-{}", id_suffix),
        change_type: first.change_type,
        sheet_name: first.sheet_name.clone(),
        cell_address: first.cell_address.clone(),
        cell_range: calculate_range(group),
        row_index: first.row_index,
        column_index: first.column_index,
        count: Some(count as i32),
        summary,
        details,
        anchor,
    }
}

fn summarize_type(t: SmlChangeType) -> &'static str {
    match t {
        SmlChangeType::CellAdded => "Added",
        SmlChangeType::CellDeleted => "Deleted",
        SmlChangeType::ValueChanged => "Value",
        SmlChangeType::FormulaChanged => "Formula",
        SmlChangeType::FormatChanged => "Format",
        _ => "Changed",
    }
}

fn calculate_range(group: &[&SmlChange]) -> Option<String> {
    if group.len() <= 1 {
        return None;
    }

    // Assuming adjacency, just take first and last
    let first = group.first().unwrap().cell_address.as_deref()?;
    let last = group.last().unwrap().cell_address.as_deref()?;

    Some(format!("{}:{}", first, last))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), (1, 1));
        assert_eq!(parse_cell_ref("B2"), (2, 2));
        assert_eq!(parse_cell_ref("AA10"), (10, 27));
    }

    #[test]
    fn test_grouping_adjacent_horizontal() {
        let mut changes = Vec::new();
        for i in 1..=3 {
            changes.push(SmlChange {
                change_type: SmlChangeType::ValueChanged,
                sheet_name: Some("Sheet1".to_string()),
                cell_address: Some(format!("{}{}", (b'A' + i - 1) as char, 1)), // A1, B1, C1
                ..Default::default()
            });
        }

        let options = SmlChangeListOptions::default();
        let items = build_change_list(&changes, &options);

        assert_eq!(items.len(), 1);
        assert_eq!(items[0].count, Some(3));
        assert_eq!(items[0].cell_range, Some("A1:C1".to_string()));
    }
}
