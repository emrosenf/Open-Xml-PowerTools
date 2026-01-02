// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

//! Internal canonical representation structures for spreadsheet comparison.
//!
//! This module provides signature structures that represent the normalized form
//! of workbooks, worksheets, and cells for comparison purposes. These signatures
//! are used internally by the SmlComparer to detect differences.

use base64::{engine::general_purpose, Engine as _};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::collections::{BTreeSet, HashMap, HashSet};

/// Represents the expanded formatting of a cell for comparison purposes.
/// Style indices are resolved to actual formatting properties.
///
/// This is a public struct (used in SmlChange) but is primarily used internally
/// for signature computation.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellFormatSignature {
    // Number format
    pub number_format_code: Option<String>,

    // Font
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_name: Option<String>,
    pub font_size: Option<f64>,
    pub font_color: Option<String>,

    // Fill
    pub fill_pattern: Option<String>,
    pub fill_foreground_color: Option<String>,
    pub fill_background_color: Option<String>,

    // Border
    pub border_left_style: Option<String>,
    pub border_left_color: Option<String>,
    pub border_right_style: Option<String>,
    pub border_right_color: Option<String>,
    pub border_top_style: Option<String>,
    pub border_top_color: Option<String>,
    pub border_bottom_style: Option<String>,
    pub border_bottom_color: Option<String>,

    // Alignment
    pub horizontal_alignment: Option<String>,
    pub vertical_alignment: Option<String>,
    pub wrap_text: bool,
    pub indent: Option<i32>,
}

impl CellFormatSignature {
    /// Returns a default cell format signature with standard Excel defaults.
    pub fn default() -> Self {
        Self {
            number_format_code: Some("General".to_string()),
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            font_name: Some("Calibri".to_string()),
            font_size: Some(11.0),
            font_color: None,
            fill_pattern: None,
            fill_foreground_color: None,
            fill_background_color: None,
            border_left_style: None,
            border_left_color: None,
            border_right_style: None,
            border_right_color: None,
            border_top_style: None,
            border_top_color: None,
            border_bottom_style: None,
            border_bottom_color: None,
            horizontal_alignment: Some("general".to_string()),
            vertical_alignment: Some("bottom".to_string()),
            wrap_text: false,
            indent: None,
        }
    }

    /// Returns a human-readable description of the differences between this format and another.
    pub fn get_difference_description(&self, other: &CellFormatSignature) -> String {
        if self == other {
            return "No difference".to_string();
        }

        let mut diffs = Vec::new();

        if self.number_format_code != other.number_format_code {
            diffs.push(format!(
                "Number format: '{:?}' → '{:?}'",
                other.number_format_code, self.number_format_code
            ));
        }
        if self.bold != other.bold {
            diffs.push(if self.bold {
                "Made bold".to_string()
            } else {
                "Removed bold".to_string()
            });
        }
        if self.italic != other.italic {
            diffs.push(if self.italic {
                "Made italic".to_string()
            } else {
                "Removed italic".to_string()
            });
        }
        if self.underline != other.underline {
            diffs.push(if self.underline {
                "Added underline".to_string()
            } else {
                "Removed underline".to_string()
            });
        }
        if self.strikethrough != other.strikethrough {
            diffs.push(if self.strikethrough {
                "Added strikethrough".to_string()
            } else {
                "Removed strikethrough".to_string()
            });
        }
        if self.font_name != other.font_name {
            diffs.push(format!(
                "Font: '{:?}' → '{:?}'",
                other.font_name, self.font_name
            ));
        }
        if self.font_size != other.font_size {
            diffs.push(format!(
                "Size: {:?} → {:?}",
                other.font_size, self.font_size
            ));
        }
        if self.font_color != other.font_color {
            diffs.push(format!(
                "Font color: {:?} → {:?}",
                other.font_color, self.font_color
            ));
        }
        if self.fill_foreground_color != other.fill_foreground_color {
            diffs.push(format!(
                "Fill color: {:?} → {:?}",
                other.fill_foreground_color, self.fill_foreground_color
            ));
        }
        if self.horizontal_alignment != other.horizontal_alignment {
            diffs.push(format!(
                "Horizontal align: {:?} → {:?}",
                other.horizontal_alignment, self.horizontal_alignment
            ));
        }
        if self.vertical_alignment != other.vertical_alignment {
            diffs.push(format!(
                "Vertical align: {:?} → {:?}",
                other.vertical_alignment, self.vertical_alignment
            ));
        }
        if self.wrap_text != other.wrap_text {
            diffs.push(if self.wrap_text {
                "Enabled wrap text".to_string()
            } else {
                "Disabled wrap text".to_string()
            });
        }

        if diffs.is_empty() {
            "Minor formatting change".to_string()
        } else {
            diffs.join("; ")
        }
    }
}

/// Internal canonical representation of a workbook for comparison.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub(crate) struct WorkbookSignature {
    pub sheets: HashMap<String, WorksheetSignature>,
    pub defined_names: HashMap<String, String>,
}

impl WorkbookSignature {
    pub fn new() -> Self {
        Self {
            sheets: HashMap::new(),
            defined_names: HashMap::new(),
        }
    }
}

impl Default for WorkbookSignature {
    fn default() -> Self {
        Self::new()
    }
}

/// Internal canonical representation of a worksheet for comparison.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub(crate) struct WorksheetSignature {
    pub name: String,
    pub relationship_id: String,
    pub cells: HashMap<String, CellSignature>,

    // Phase 2: Row-level data for alignment
    pub populated_rows: BTreeSet<i32>,
    pub populated_columns: BTreeSet<i32>,
    pub row_signatures: HashMap<i32, String>,
    pub column_signatures: HashMap<i32, String>,

    // Phase 3: Comments, data validation, merged cells, hyperlinks
    pub comments: HashMap<String, CommentSignature>,
    pub data_validations: HashMap<String, DataValidationSignature>,
    pub merged_cell_ranges: HashSet<String>,
    pub hyperlinks: HashMap<String, HyperlinkSignature>,
}

impl WorksheetSignature {
    pub fn new(name: String, relationship_id: String) -> Self {
        Self {
            name,
            relationship_id,
            cells: HashMap::new(),
            populated_rows: BTreeSet::new(),
            populated_columns: BTreeSet::new(),
            row_signatures: HashMap::new(),
            column_signatures: HashMap::new(),
            comments: HashMap::new(),
            data_validations: HashMap::new(),
            merged_cell_ranges: HashSet::new(),
            hyperlinks: HashMap::new(),
        }
    }

    /// Get all cells in a specific row.
    pub fn get_cells_in_row(&self, row: i32) -> Vec<&CellSignature> {
        let mut cells: Vec<&CellSignature> = self.cells.values().filter(|c| c.row == row).collect();
        cells.sort_by_key(|c| c.column);
        cells
    }

    /// Get all cells in a specific column.
    pub fn get_cells_in_column(&self, col: i32) -> Vec<&CellSignature> {
        let mut cells: Vec<&CellSignature> =
            self.cells.values().filter(|c| c.column == col).collect();
        cells.sort_by_key(|c| c.row);
        cells
    }

    /// Compute a content hash representing this sheet's overall content (for rename detection).
    /// Uses SHA256 for consistency with C# implementation.
    pub fn compute_content_hash(&self) -> String {
        let mut cells: Vec<&CellSignature> = self.cells.values().collect();
        cells.sort_by_key(|c| (c.row, c.column));

        let mut content = String::new();
        for cell in cells {
            content.push_str(&format!(
                "{}:{}|",
                cell.address,
                cell.resolved_value.as_deref().unwrap_or("")
            ));
        }

        let mut hasher = Sha256::new();
        hasher.update(content.as_bytes());
        let result = hasher.finalize();
        general_purpose::STANDARD.encode(&result)
    }
}

/// Internal canonical representation of a cell for comparison.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub(crate) struct CellSignature {
    pub address: String,
    pub row: i32,
    pub column: i32,
    pub resolved_value: Option<String>,
    pub formula: Option<String>,
    pub content_hash: String,
    pub format: CellFormatSignature,
}

impl CellSignature {
    /// Computes a content hash for cell comparison.
    /// Uses SHA256 for consistency with C# implementation.
    pub fn compute_hash(value: Option<&str>, formula: Option<&str>) -> String {
        let content = format!("{}|{}", value.unwrap_or(""), formula.unwrap_or(""));
        let mut hasher = Sha256::new();
        hasher.update(content.as_bytes());
        let result = hasher.finalize();
        general_purpose::STANDARD.encode(&result)
    }
}

/// Phase 3: Represents a cell comment for comparison.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub(crate) struct CommentSignature {
    pub cell_address: String,
    pub author: String,
    pub text: String,
}

impl CommentSignature {
    #[allow(dead_code)]
    pub fn compute_hash(&self) -> String {
        let content = format!("{}|{}", self.author, self.text);
        let mut hasher = Sha256::new();
        hasher.update(content.as_bytes());
        let result = hasher.finalize();
        general_purpose::STANDARD.encode(&result)
    }
}

/// Phase 3: Represents a data validation rule for comparison.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub(crate) struct DataValidationSignature {
    pub cell_range: String,
    /// Type: list, whole, decimal, date, time, textLength, custom
    pub validation_type: String,
    /// Operator: between, notBetween, equal, notEqual, etc.
    pub operator: Option<String>,
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    pub allow_blank: bool,
    pub show_drop_down: bool,
    pub show_input_message: bool,
    pub show_error_message: bool,
    pub error_title: Option<String>,
    pub error: Option<String>,
    pub prompt_title: Option<String>,
    pub prompt: Option<String>,
}

impl DataValidationSignature {
    pub fn compute_hash(&self) -> String {
        let content = format!(
            "{}|{}|{}|{}|{}|{}",
            self.validation_type,
            self.operator.as_deref().unwrap_or(""),
            self.formula1.as_deref().unwrap_or(""),
            self.formula2.as_deref().unwrap_or(""),
            self.allow_blank,
            self.show_drop_down
        );
        let mut hasher = Sha256::new();
        hasher.update(content.as_bytes());
        let result = hasher.finalize();
        general_purpose::STANDARD.encode(&result)
    }

    pub fn to_string(&self) -> String {
        let mut parts = vec![format!("Type: {}", self.validation_type)];
        if let Some(ref op) = self.operator {
            parts.push(format!("Operator: {}", op));
        }
        if let Some(ref f1) = self.formula1 {
            parts.push(format!("Formula1: {}", f1));
        }
        if let Some(ref f2) = self.formula2 {
            parts.push(format!("Formula2: {}", f2));
        }
        parts.join(", ")
    }
}

/// Phase 3: Represents a hyperlink for comparison.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub(crate) struct HyperlinkSignature {
    pub cell_address: String,
    pub target: String,
    pub display: Option<String>,
    pub tooltip: Option<String>,
}

impl HyperlinkSignature {
    pub fn compute_hash(&self) -> String {
        let content = format!("{}|{}", self.target, self.display.as_deref().unwrap_or(""));
        let mut hasher = Sha256::new();
        hasher.update(content.as_bytes());
        let result = hasher.finalize();
        general_purpose::STANDARD.encode(&result)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cell_format_default_has_expected_values() {
        let format = CellFormatSignature::default();
        assert_eq!(format.number_format_code, Some("General".to_string()));
        assert!(!format.bold);
        assert!(!format.italic);
        assert_eq!(format.font_name, Some("Calibri".to_string()));
        assert_eq!(format.font_size, Some(11.0));
        assert_eq!(format.horizontal_alignment, Some("general".to_string()));
        assert_eq!(format.vertical_alignment, Some("bottom".to_string()));
        assert!(!format.wrap_text);
    }

    #[test]
    fn cell_format_equality_works() {
        let format1 = CellFormatSignature::default();
        let format2 = CellFormatSignature::default();
        assert_eq!(format1, format2);
    }

    #[test]
    fn cell_format_difference_description() {
        let format1 = CellFormatSignature::default();
        let mut format2 = CellFormatSignature::default();
        format2.bold = true;

        let desc = format2.get_difference_description(&format1);
        assert!(desc.contains("Made bold"));
    }

    #[test]
    fn cell_signature_compute_hash() {
        let hash1 = CellSignature::compute_hash(Some("value"), Some("=A1+B1"));
        let hash2 = CellSignature::compute_hash(Some("value"), Some("=A1+B1"));
        let hash3 = CellSignature::compute_hash(Some("different"), Some("=A1+B1"));

        assert_eq!(hash1, hash2);
        assert_ne!(hash1, hash3);
    }

    #[test]
    fn workbook_signature_creates_empty() {
        let sig = WorkbookSignature::new();
        assert!(sig.sheets.is_empty());
        assert!(sig.defined_names.is_empty());
    }

    #[test]
    fn worksheet_signature_creates_with_name() {
        let sig = WorksheetSignature::new("Sheet1".to_string(), "rId1".to_string());
        assert_eq!(sig.name, "Sheet1");
        assert_eq!(sig.relationship_id, "rId1");
        assert!(sig.cells.is_empty());
    }

    #[test]
    fn worksheet_compute_content_hash() {
        let mut sig = WorksheetSignature::new("Sheet1".to_string(), "rId1".to_string());

        sig.cells.insert(
            "A1".to_string(),
            CellSignature {
                address: "A1".to_string(),
                row: 1,
                column: 1,
                resolved_value: Some("test".to_string()),
                formula: None,
                content_hash: String::new(),
                format: CellFormatSignature::default(),
            },
        );

        let hash = sig.compute_content_hash();
        assert!(!hash.is_empty());
    }

    #[test]
    fn comment_signature_compute_hash() {
        let comment = CommentSignature {
            cell_address: "A1".to_string(),
            author: "Test Author".to_string(),
            text: "Test comment".to_string(),
        };

        let hash = comment.compute_hash();
        assert!(!hash.is_empty());
    }

    #[test]
    fn data_validation_signature_to_string() {
        let dv = DataValidationSignature {
            cell_range: "A1:A10".to_string(),
            validation_type: "list".to_string(),
            operator: Some("between".to_string()),
            formula1: Some("1".to_string()),
            formula2: Some("10".to_string()),
            allow_blank: true,
            show_drop_down: true,
            show_input_message: false,
            show_error_message: false,
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
        };

        let s = dv.to_string();
        assert!(s.contains("Type: list"));
        assert!(s.contains("Operator: between"));
        assert!(s.contains("Formula1: 1"));
    }

    #[test]
    fn hyperlink_signature_compute_hash() {
        let hyperlink = HyperlinkSignature {
            cell_address: "A1".to_string(),
            target: "https://example.com".to_string(),
            display: Some("Example".to_string()),
            tooltip: Some("Click here".to_string()),
        };

        let hash = hyperlink.compute_hash();
        assert!(!hash.is_empty());
    }
}
