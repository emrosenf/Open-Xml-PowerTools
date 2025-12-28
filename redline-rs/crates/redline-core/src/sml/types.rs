// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

//! Public types for SmlComparer results and changes.
//!
//! This module defines SmlChange and SmlChangeType which are used to represent
//! individual changes detected during spreadsheet comparison.

use serde::{Deserialize, Serialize};
use crate::sml::CellFormatSignature;

/// Types of changes detected during spreadsheet comparison.
/// 100% parity with C# SmlChangeType enum.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub enum SmlChangeType {
    // Workbook structure
    SheetAdded,
    SheetDeleted,
    SheetRenamed,

    // Row/column structure (Phase 2)
    RowInserted,
    RowDeleted,
    ColumnInserted,
    ColumnDeleted,

    // Cell content
    CellAdded,
    CellDeleted,
    ValueChanged,
    FormulaChanged,
    FormatChanged,

    // Phase 3: Named ranges
    NamedRangeAdded,
    NamedRangeDeleted,
    NamedRangeChanged,

    // Phase 3: Comments
    CommentAdded,
    CommentDeleted,
    CommentChanged,

    // Phase 3: Data validation
    DataValidationAdded,
    DataValidationDeleted,
    DataValidationChanged,

    // Phase 3: Merged cells
    MergedCellAdded,
    MergedCellDeleted,

    // Phase 3: Conditional formatting
    ConditionalFormatAdded,
    ConditionalFormatDeleted,
    ConditionalFormatChanged,

    // Phase 3: Hyperlinks
    HyperlinkAdded,
    HyperlinkDeleted,
    HyperlinkChanged,
}

/// Represents a single change between two spreadsheets.
/// 100% parity with C# SmlChange class.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlChange {
    pub change_type: SmlChangeType,
    pub sheet_name: Option<String>,
    pub cell_address: Option<String>,

    // Phase 2: Row/column indices for structural changes
    pub row_index: Option<i32>,
    pub column_index: Option<i32>,

    // Phase 2: For sheet rename detection
    pub old_sheet_name: Option<String>,

    pub old_value: Option<String>,
    pub new_value: Option<String>,
    pub old_formula: Option<String>,
    pub new_formula: Option<String>,
    pub old_format: Option<CellFormatSignature>,
    pub new_format: Option<CellFormatSignature>,

    // Phase 3: Named range properties
    pub named_range_name: Option<String>,
    pub old_named_range_value: Option<String>,
    pub new_named_range_value: Option<String>,

    // Phase 3: Comment properties
    pub old_comment: Option<String>,
    pub new_comment: Option<String>,
    pub comment_author: Option<String>,

    // Phase 3: Data validation properties
    pub data_validation_type: Option<String>,
    pub old_data_validation: Option<String>,
    pub new_data_validation: Option<String>,

    // Phase 3: Merged cell range
    pub merged_cell_range: Option<String>,

    // Phase 3: Hyperlink properties
    pub old_hyperlink: Option<String>,
    pub new_hyperlink: Option<String>,

    // Phase 3: Conditional formatting properties
    pub conditional_format_range: Option<String>,
    pub old_conditional_format: Option<String>,
    pub new_conditional_format: Option<String>,
}

impl SmlChange {
    /// Returns a human-readable description of this change.
    /// 100% parity with C# GetDescription() method.
    pub fn get_description(&self) -> String {
        match self.change_type {
            SmlChangeType::SheetAdded => {
                format!("Sheet '{}' was added", self.sheet_name.as_deref().unwrap_or(""))
            }
            SmlChangeType::SheetDeleted => {
                format!("Sheet '{}' was deleted", self.sheet_name.as_deref().unwrap_or(""))
            }
            SmlChangeType::SheetRenamed => {
                format!(
                    "Sheet '{}' was renamed to '{}'",
                    self.old_sheet_name.as_deref().unwrap_or(""),
                    self.sheet_name.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::RowInserted => {
                format!(
                    "Row {} was inserted in sheet '{}'",
                    self.row_index.unwrap_or(0),
                    self.sheet_name.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::RowDeleted => {
                format!(
                    "Row {} was deleted from sheet '{}'",
                    self.row_index.unwrap_or(0),
                    self.sheet_name.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::ColumnInserted => {
                format!(
                    "Column {} was inserted in sheet '{}'",
                    get_column_letter(self.column_index.unwrap_or(0)),
                    self.sheet_name.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::ColumnDeleted => {
                format!(
                    "Column {} was deleted from sheet '{}'",
                    get_column_letter(self.column_index.unwrap_or(0)),
                    self.sheet_name.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::CellAdded => {
                format!(
                    "Cell {}!{} was added with value '{}'",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or(""),
                    self.new_value.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::CellDeleted => {
                format!(
                    "Cell {}!{} was deleted (had value '{}')",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or(""),
                    self.old_value.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::ValueChanged => {
                format!(
                    "Cell {}!{} value changed from '{}' to '{}'",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or(""),
                    self.old_value.as_deref().unwrap_or(""),
                    self.new_value.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::FormulaChanged => {
                format!(
                    "Cell {}!{} formula changed from '{}' to '{}'",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or(""),
                    self.old_formula.as_deref().unwrap_or(""),
                    self.new_formula.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::FormatChanged => {
                format!(
                    "Cell {}!{} formatting changed",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or("")
                )
            }

            // Phase 3: Named ranges
            SmlChangeType::NamedRangeAdded => {
                format!(
                    "Named range '{}' was added with value '{}'",
                    self.named_range_name.as_deref().unwrap_or(""),
                    self.new_named_range_value.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::NamedRangeDeleted => {
                format!(
                    "Named range '{}' was deleted (had value '{}')",
                    self.named_range_name.as_deref().unwrap_or(""),
                    self.old_named_range_value.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::NamedRangeChanged => {
                format!(
                    "Named range '{}' changed from '{}' to '{}'",
                    self.named_range_name.as_deref().unwrap_or(""),
                    self.old_named_range_value.as_deref().unwrap_or(""),
                    self.new_named_range_value.as_deref().unwrap_or("")
                )
            }

            // Phase 3: Comments
            SmlChangeType::CommentAdded => {
                format!(
                    "Comment added to {}!{}: '{}'",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or(""),
                    truncate_for_display(self.new_comment.as_deref().unwrap_or(""), 50)
                )
            }
            SmlChangeType::CommentDeleted => {
                format!(
                    "Comment deleted from {}!{}",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::CommentChanged => {
                format!(
                    "Comment changed at {}!{}",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or("")
                )
            }

            // Phase 3: Data validation
            SmlChangeType::DataValidationAdded => {
                format!(
                    "Data validation ({}) added to {}!{}",
                    self.data_validation_type.as_deref().unwrap_or(""),
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::DataValidationDeleted => {
                format!(
                    "Data validation removed from {}!{}",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::DataValidationChanged => {
                format!(
                    "Data validation changed at {}!{}",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or("")
                )
            }

            // Phase 3: Merged cells
            SmlChangeType::MergedCellAdded => {
                format!(
                    "Merged cell region {} added in sheet '{}'",
                    self.merged_cell_range.as_deref().unwrap_or(""),
                    self.sheet_name.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::MergedCellDeleted => {
                format!(
                    "Merged cell region {} removed from sheet '{}'",
                    self.merged_cell_range.as_deref().unwrap_or(""),
                    self.sheet_name.as_deref().unwrap_or("")
                )
            }

            // Phase 3: Conditional formatting
            SmlChangeType::ConditionalFormatAdded => {
                format!(
                    "Conditional formatting added to {} in sheet '{}'",
                    self.conditional_format_range.as_deref().unwrap_or(""),
                    self.sheet_name.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::ConditionalFormatDeleted => {
                format!(
                    "Conditional formatting removed from {} in sheet '{}'",
                    self.conditional_format_range.as_deref().unwrap_or(""),
                    self.sheet_name.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::ConditionalFormatChanged => {
                format!(
                    "Conditional formatting changed at {} in sheet '{}'",
                    self.conditional_format_range.as_deref().unwrap_or(""),
                    self.sheet_name.as_deref().unwrap_or("")
                )
            }

            // Phase 3: Hyperlinks
            SmlChangeType::HyperlinkAdded => {
                format!(
                    "Hyperlink added to {}!{}: '{}'",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or(""),
                    self.new_hyperlink.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::HyperlinkDeleted => {
                format!(
                    "Hyperlink removed from {}!{}",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or("")
                )
            }
            SmlChangeType::HyperlinkChanged => {
                format!(
                    "Hyperlink changed at {}!{} from '{}' to '{}'",
                    self.sheet_name.as_deref().unwrap_or(""),
                    self.cell_address.as_deref().unwrap_or(""),
                    self.old_hyperlink.as_deref().unwrap_or(""),
                    self.new_hyperlink.as_deref().unwrap_or("")
                )
            }
        }
    }
}

impl Default for SmlChange {
    fn default() -> Self {
        Self {
            change_type: SmlChangeType::ValueChanged,
            sheet_name: None,
            cell_address: None,
            row_index: None,
            column_index: None,
            old_sheet_name: None,
            old_value: None,
            new_value: None,
            old_formula: None,
            new_formula: None,
            old_format: None,
            new_format: None,
            named_range_name: None,
            old_named_range_value: None,
            new_named_range_value: None,
            old_comment: None,
            new_comment: None,
            comment_author: None,
            data_validation_type: None,
            old_data_validation: None,
            new_data_validation: None,
            merged_cell_range: None,
            old_hyperlink: None,
            new_hyperlink: None,
            conditional_format_range: None,
            old_conditional_format: None,
            new_conditional_format: None,
        }
    }
}



// Helper functions

fn get_column_letter(column_number: i32) -> String {
    let mut result = String::new();
    let mut num = column_number;
    
    while num > 0 {
        num -= 1;
        result.insert(0, (b'A' + (num % 26) as u8) as char);
        num /= 26;
    }
    
    result
}

fn truncate_for_display(text: &str, max_length: usize) -> String {
    if text.is_empty() {
        return String::new();
    }
    if text.len() <= max_length {
        return text.to_string();
    }
    format!("{}...", &text[..max_length - 3])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn change_type_serializes_correctly() {
        let change_type = SmlChangeType::SheetAdded;
        let json = serde_json::to_string(&change_type).unwrap();
        assert_eq!(json, "\"SheetAdded\"");
    }

    #[test]
    fn sml_change_get_description_works() {
        let mut change = SmlChange::default();
        change.change_type = SmlChangeType::SheetAdded;
        change.sheet_name = Some("Sheet1".to_string());

        let desc = change.get_description();
        assert_eq!(desc, "Sheet 'Sheet1' was added");
    }

    #[test]
    fn sml_comparison_result_computes_statistics() {
        let mut result = SmlComparisonResult::new();
        
        result.changes.push(SmlChange {
            change_type: SmlChangeType::ValueChanged,
            ..Default::default()
        });
        result.changes.push(SmlChange {
            change_type: SmlChangeType::FormulaChanged,
            ..Default::default()
        });
        result.changes.push(SmlChange {
            change_type: SmlChangeType::ValueChanged,
            ..Default::default()
        });

        assert_eq!(result.total_changes(), 3);
        assert_eq!(result.value_changes(), 2);
        assert_eq!(result.formula_changes(), 1);
        assert_eq!(result.format_changes(), 0);
    }



    #[test]
    fn get_column_letter_works() {
        assert_eq!(get_column_letter(1), "A");
        assert_eq!(get_column_letter(26), "Z");
        assert_eq!(get_column_letter(27), "AA");
        assert_eq!(get_column_letter(52), "AZ");
    }

    #[test]
    fn truncate_for_display_works() {
        assert_eq!(truncate_for_display("", 10), "");
        assert_eq!(truncate_for_display("short", 10), "short");
        assert_eq!(
            truncate_for_display("This is a very long text", 10),
            "This is..."
        );
    }


}
