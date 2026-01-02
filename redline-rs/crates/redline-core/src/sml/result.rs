use crate::sml::types::{SmlChange, SmlChangeType};
use serde::{Deserialize, Serialize};

/// Result of comparing two spreadsheets, containing all detected changes.
/// 100% parity with C# SmlComparisonResult class.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlComparisonResult {
    pub changes: Vec<SmlChange>,
}

impl SmlComparisonResult {
    pub fn new() -> Self {
        Self {
            changes: Vec::new(),
        }
    }

    /// Add a change to the result.
    pub fn add_change(&mut self, change: SmlChange) {
        self.changes.push(change);
    }

    /// Get total number of changes.
    pub fn total_changes(&self) -> usize {
        self.changes.len()
    }

    /// Get number of value changes.
    pub fn value_changes(&self) -> usize {
        self.changes
            .iter()
            .filter(|c| c.change_type == SmlChangeType::ValueChanged)
            .count()
    }

    /// Get number of formula changes.
    pub fn formula_changes(&self) -> usize {
        self.changes
            .iter()
            .filter(|c| c.change_type == SmlChangeType::FormulaChanged)
            .count()
    }

    /// Get number of format changes.
    pub fn format_changes(&self) -> usize {
        self.changes
            .iter()
            .filter(|c| c.change_type == SmlChangeType::FormatChanged)
            .count()
    }

    /// Get number of cells added.
    pub fn cells_added(&self) -> usize {
        self.changes
            .iter()
            .filter(|c| c.change_type == SmlChangeType::CellAdded)
            .count()
    }

    /// Get number of cells deleted.
    pub fn cells_deleted(&self) -> usize {
        self.changes
            .iter()
            .filter(|c| c.change_type == SmlChangeType::CellDeleted)
            .count()
    }

    /// Get number of sheets added.
    pub fn sheets_added(&self) -> usize {
        self.changes
            .iter()
            .filter(|c| c.change_type == SmlChangeType::SheetAdded)
            .count()
    }

    /// Get number of sheets deleted.
    pub fn sheets_deleted(&self) -> usize {
        self.changes
            .iter()
            .filter(|c| c.change_type == SmlChangeType::SheetDeleted)
            .count()
    }

    /// Get number of sheets renamed (Phase 2).
    pub fn sheets_renamed(&self) -> usize {
        self.changes
            .iter()
            .filter(|c| c.change_type == SmlChangeType::SheetRenamed)
            .count()
    }

    /// Get number of rows inserted (Phase 2).
    pub fn rows_inserted(&self) -> usize {
        self.changes
            .iter()
            .filter(|c| c.change_type == SmlChangeType::RowInserted)
            .count()
    }

    /// Get number of rows deleted (Phase 2).
    pub fn rows_deleted(&self) -> usize {
        self.changes
            .iter()
            .filter(|c| c.change_type == SmlChangeType::RowDeleted)
            .count()
    }

    /// Get number of columns inserted (Phase 2).
    pub fn columns_inserted(&self) -> usize {
        self.changes
            .iter()
            .filter(|c| c.change_type == SmlChangeType::ColumnInserted)
            .count()
    }

    /// Get number of columns deleted (Phase 2).
    pub fn columns_deleted(&self) -> usize {
        self.changes
            .iter()
            .filter(|c| c.change_type == SmlChangeType::ColumnDeleted)
            .count()
    }

    /// Get all changes for a specific sheet.
    pub fn get_changes_by_sheet(&self, sheet_name: &str) -> Vec<&SmlChange> {
        self.changes
            .iter()
            .filter(|c| c.sheet_name.as_deref() == Some(sheet_name))
            .collect()
    }

    /// Get all changes of a specific type.
    pub fn get_changes_by_type(&self, change_type: SmlChangeType) -> Vec<&SmlChange> {
        self.changes
            .iter()
            .filter(|c| c.change_type == change_type)
            .collect()
    }

    /// Export the comparison result to JSON.
    pub fn to_json(&self) -> String {
        serde_json::to_string_pretty(self).unwrap_or_default()
    }
}

impl Default for SmlComparisonResult {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetComparisonResult {
    pub name: String,
    pub status: SheetStatus,
    pub cell_changes: Vec<CellChange>,
    pub row_changes: Vec<RowChange>,
    pub column_changes: Vec<ColumnChange>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum SheetStatus {
    Added,
    Deleted,
    Modified,
    Unchanged,
    Renamed,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellChange {
    pub cell_address: String,
    pub change_type: CellChangeType,
    pub old_value: Option<String>,
    pub new_value: Option<String>,
    pub old_formula: Option<String>,
    pub new_formula: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum CellChangeType {
    Added,
    Deleted,
    ValueChanged,
    FormulaChanged,
    FormatChanged,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RowChange {
    pub row_number: u32,
    pub change_type: RowChangeType,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum RowChangeType {
    Added,
    Deleted,
    Modified,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColumnChange {
    pub column_letter: String,
    pub change_type: ColumnChangeType,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum ColumnChangeType {
    Added,
    Deleted,
    Modified,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn result_serializes_to_json() {
        let result = SmlComparisonResult::new();
        let json = result.to_json();
        // The JSON contains the "changes" field from the struct
        assert!(
            json.contains("\"changes\""),
            "JSON should contain 'changes' field"
        );
    }
}
